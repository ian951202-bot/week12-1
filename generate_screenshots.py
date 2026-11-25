import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyBboxPatch, Rectangle, Arrow, ConnectionPatch
import numpy as np
import os

# Set font for Traditional Chinese
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'Segoe UI', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

os.makedirs("images", exist_ok=True)

# Helper to draw rounded box
def add_card(ax, xy, width, height, box_color, border_color='none', radius=0.02, alpha=1.0):
    box = FancyBboxPatch(xy, width, height,
                         boxstyle=f"round,pad=0.01,rounding_size={radius}",
                         facecolor=box_color, edgecolor=border_color, linewidth=1.5, alpha=alpha)
    ax.add_patch(box)
    return box

# ========================================================
# Figure 1: Excel Interface Overview
# ========================================================
def create_fig1():
    fig, ax = plt.subplots(figsize=(12, 7.2), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    # Outer Excel Window
    add_card(ax, (1, 1), 98, 98, "#F3F2F1", "#A6A6A6", radius=0.5)
    
    # Title Bar (Green Excel Theme)
    add_card(ax, (1, 91), 98, 8, "#107C41", "#107C41", radius=0.3)
    ax.text(3, 95, "AutoSave [ON]  |  Shipping_Demo_30.xlsx - Excel", color="white", fontsize=11, fontweight='bold', va='center')
    ax.text(78, 95, "搜尋 (Alt+Q)         _   [ ]   X", color="white", fontsize=11, va='center')
    
    # Ribbon Tabs Bar
    add_card(ax, (1, 84), 98, 7, "#FFFFFF", "#E1DFDD", radius=0.1)
    tabs = ["檔案", "常用 (Home)", "插入", "繪圖", "版面配置", "公式 (Formulas)", "資料 (Data)", "校閱", "檢視", "說明"]
    x_pos = 3
    for i, tab in enumerate(tabs):
        if "常用" in tab:
            add_card(ax, (x_pos-1, 84), len(tab)*1.6 + 2, 7, "#E6F2EA", "#107C41", radius=0.1)
            ax.text(x_pos, 87.5, tab, color="#107C41", fontsize=10, fontweight='bold', va='center')
        elif "資料" in tab or "公式" in tab:
            ax.text(x_pos, 87.5, tab, color="#005A9E", fontsize=10, fontweight='bold', va='center')
        else:
            ax.text(x_pos, 87.5, tab, color="#323130", fontsize=10, va='center')
        x_pos += len(tab)*1.6 + 3
        
    # Ribbon Command Groups
    add_card(ax, (1, 68), 98, 16, "#F8F9FA", "#D2D0CE", radius=0.1)
    
    # Groups inside ribbon
    # Group 1: 剪貼簿
    add_card(ax, (2, 70), 12, 13, "#FFFFFF", "#E1DFDD", radius=0.2)
    ax.text(8, 79, "貼上 / 複製\n剪下 / 格式刷", ha='center', va='center', fontsize=8.5, color="#323130")
    ax.text(8, 71.5, "剪貼簿", ha='center', va='center', fontsize=8, color="#605E5C", fontweight='bold')
    
    # Group 2: 字型
    add_card(ax, (15, 70), 17, 13, "#FFFFFF", "#E1DFDD", radius=0.2)
    ax.text(23.5, 79, "微軟正黑體  11  ▼\n[B] [I] [U] [田] [色]", ha='center', va='center', fontsize=8.5, color="#323130")
    ax.text(23.5, 71.5, "字型", ha='center', va='center', fontsize=8, color="#605E5C", fontweight='bold')
    
    # Group 3: 對齊方式
    add_card(ax, (33, 70), 16, 13, "#FFFFFF", "#E1DFDD", radius=0.2)
    ax.text(41, 79, "靠左 / 置中 / 靠右\n跨欄置中 / 自動換行", ha='center', va='center', fontsize=8.5, color="#323130")
    ax.text(41, 71.5, "對齊方式", ha='center', va='center', fontsize=8, color="#605E5C", fontweight='bold')
    
    # Group 4: 數值
    add_card(ax, (50, 70), 16, 13, "#FFFFFF", "#E1DFDD", radius=0.2)
    ax.text(58, 79, "通用格式  ▼\n$  %  ,  .00  .0", ha='center', va='center', fontsize=8.5, color="#323130")
    ax.text(58, 71.5, "數值", ha='center', va='center', fontsize=8, color="#605E5C", fontweight='bold')
    
    # Group 5: 編輯與資料工具
    add_card(ax, (67, 70), 31, 13, "#FFFFFF", "#E1DFDD", radius=0.2)
    ax.text(82.5, 79, "自動加總 ∑  |  填滿 (快速填入 Ctrl+E)  |  清除\n排序與篩選 ▽  |  尋找與選取 ", ha='center', va='center', fontsize=8.5, color="#323130")
    ax.text(82.5, 71.5, "編輯與分析工具", ha='center', va='center', fontsize=8, color="#605E5C", fontweight='bold')
    
    # Formula Bar & Name Box Area
    add_card(ax, (1, 61), 98, 6.5, "#FFFFFF", "#E1DFDD", radius=0.1)
    # Name Box
    add_card(ax, (2, 62), 10, 4.5, "#F3F2F1", "#8A8886", radius=0.1)
    ax.text(7, 64.2, "B2", ha='center', va='center', fontsize=10, fontweight='bold', color="#201F1E")
    ax.text(13.5, 64.2, "X  V  fx", ha='center', va='center', fontsize=10, color="#605E5C")
    # Formula Input Box
    add_card(ax, (17, 62), 81, 4.5, "#FFFFFF", "#8A8886", radius=0.1)
    ax.text(18.5, 64.2, "=TRIM(B2)   [資料編輯列：輸入/檢視儲存格公式或文字]", va='center', fontsize=9.5, color="#004578", fontfamily='Microsoft JhengHei')
    
    # Spreadsheet Grid Area
    add_card(ax, (1, 9), 98, 51.5, "#FFFFFF", "#D2D0CE", radius=0.1)
    # Column Headers (A to G)
    cols = ["", "A (項次)", "B (提單號碼)", "C (船名航次)", "D (裝貨港)", "E (卸貨港)", "F (貨櫃號碼)", "G (海運費)"]
    widths = [4, 8, 16, 21, 11, 11, 15, 12]
    cur_x = 2
    for idx, (c_name, w) in enumerate(zip(cols, widths)):
        add_card(ax, (cur_x, 56), w, 4, "#E1DFDD" if idx>0 else "#C8C6C4", "#A6A6A6", radius=0.0)
        ax.text(cur_x + w/2, 58, c_name, ha='center', va='center', fontsize=8.5, fontweight='bold', color="#323130")
        cur_x += w
        
    # Grid Data Rows (Row 1 to Row 8)
    sample_rows = [
        ["1", "1", "EMC-2026-001  ", "EVER GIVEN / 0142-052W", "twkhh", "USLAX", "EMCU1234567 ", "$2,450.00"],
        ["2", "2", "WHL-2026-088", "WAN HAI 315 / 0089-012N", "TWKEL", "cnsHa", "WHLU9876543", "USD 1200"],
        ["3", "3", "YML-2026-105 ", "YM WELLBEING / 0214-088E", "twkhh", "NLRTM", "YMLU4567890", "EUR 3,100"],
        ["4", "4", "OOCL-2026-302", "OOCL POLAND / 0055-031W", "TWKHH", "USLGB", "OOLU2345678", "$2,600"],
        ["5", "5", "MSK-2026-991", "MAERSK MC-KINNEY / 0411-002W", "TWKEL", "DEHAM", "MSKU3456789", "3400 USD"],
        ["6", "6", "CMA-2026-412", "CMA CGM ANTOINE / 0918-077E", "twkhh", "FRLEH", "CMAU8765432", "$1,450"],
        ["7", "7", "ONE-2026-550", "ONE APUS / 0022-019E", "TWKEL", "USOAK", "ONEU5678901", "2850"]
    ]
    
    y_grid = 51.5
    for r_idx, r_data in enumerate(sample_rows):
        cur_x = 2
        bg = "#F9F9FB" if r_idx % 2 == 1 else "#FFFFFF"
        # highlight cell B2
        for c_idx, (val, w) in enumerate(zip(r_data, widths)):
            cell_bg = bg
            border_c = "#E1DFDD"
            if r_idx == 0 and c_idx == 2: # Cell B2
                cell_bg = "#C7E0F4" # Selected cell highlight
                border_c = "#0078D4"
            add_card(ax, (cur_x, y_grid), w, 4.5, cell_bg, border_c, radius=0.0)
            align = 'center' if c_idx in [0, 1, 4, 5] else ('right' if c_idx == 7 else 'left')
            t_x = cur_x + w/2 if align=='center' else (cur_x + w - 1 if align=='right' else cur_x + 1)
            ax.text(t_x, y_grid + 2.25, val, ha=align, va='center', fontsize=8, color="#201F1E")
            cur_x += w
        y_grid -= 4.5
        
    # Sheet Tabs & Status Bar
    add_card(ax, (1, 4.5), 98, 4.5, "#F3F2F1", "#D2D0CE", radius=0.1)
    add_card(ax, (6, 5), 22, 3.5, "#FFFFFF", "#107C41", radius=0.1)
    ax.text(17, 6.7, "航運原始資料_未清洗 (Raw)", ha='center', va='center', fontsize=8.5, fontweight='bold', color="#107C41")
    ax.text(32, 6.7, "+ 航運標準資料 (Cleaned)   +", va='center', fontsize=8.5, color="#605E5C")
    
    # Status bar
    add_card(ax, (1, 1), 98, 3.5, "#107C41", "#107C41", radius=0.1)
    ax.text(3, 2.7, "就緒 (Ready)   |   協助協助工具: 良好", color="white", fontsize=8.5, va='center')
    ax.text(82, 2.7, "平均值: 2,350 | 計數: 7 | 100% ▰▰▰", color="white", fontsize=8.5, va='center')
    
    # Callout Annotations pointing to interface elements
    callouts = [
        ((7, 95), (14, 98), "① 快速存取與標題列\n(檔案儲存/雲端同動)"),
        ((35, 87), (35, 93), "② 功能區索引標籤 (Ribbon Tabs)\n常用/公式/資料是最核心操作區"),
        ((75, 78), (75, 88), "③ 功能群組與指令按鈕\n(快速填入/排序篩選/資料工具)"),
        ((6, 64), (6, 58), "④ 名稱方塊\n(顯示目前選取儲存格位址)"),
        ((35, 64), (35, 59), "⑤ 資料編輯列 (Formula Bar)\n檢視與編輯儲存格公式核心"),
        ((18, 53), (18, 45), "⑥ 儲存格網格與工作區\n欄位(A,B,C)與列號(1,2,3)交會處"),
        ((15, 6), (15, 0), "⑦ 工作表標籤 (Sheet Tabs)\n切換原始資料與清洗完成報表"),
        ((85, 2.5), (85, -3), "⑧ 狀態列 (Status Bar)\n即時顯示選取區域之計數與平均值")
    ]
    
    for pt_from, pt_to, text in callouts:
        ax.annotate(text, xy=pt_from, xytext=pt_to,
                    arrowprops=dict(facecolor='#D83B01', edgecolor='#A80000', width=1.5, headwidth=6, headlength=6, shrink=0.08),
                    bbox=dict(boxstyle="round,pad=0.3", fc="#FFF4CE", ec="#D83B01", lw=1.2),
                    fontsize=8, fontweight='bold', color="#7A0000", ha='center', va='center')
                    
    plt.tight_layout()
    plt.savefig("images/fig1_excel_interface.png", dpi=200, bbox_inches='tight')
    plt.close()
    print("Generated fig1_excel_interface.png")

# ========================================================
# Figure 2: Bad Table vs Good Table (Data Hygiene)
# ========================================================
def create_fig2():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 6.5), dpi=200)
    
    for ax in [ax1, ax2]:
        ax.set_xlim(0, 100)
        ax.set_ylim(0, 100)
        ax.axis('off')
        
    # LEFT: BAD TABLE (報表型視覺排版 - 禁用於分析)
    add_card(ax1, (2, 2), 96, 96, "#FFF9F9", "#E81123", radius=0.4)
    ax1.text(50, 94, "[錯誤] 不合格的資料結構 (報表排版型 - 無法有效分析)", ha='center', va='center', fontsize=11, fontweight='bold', color="#A80000")
    
    # Merged title
    add_card(ax1, (5, 84), 90, 6, "#E81123", "#A80000", radius=0.1)
    ax1.text(50, 87, "2026年第4季 美國線航運提單與運費彙總報表 (大標題合併儲存格)", ha='center', va='center', fontsize=9, fontweight='bold', color="white")
    
    # Double Header with Merged Cells
    add_card(ax1, (5, 76), 45, 7, "#FDE7E9", "#E81123", radius=0.0)
    ax1.text(27.5, 79.5, "基本航運航程資訊 (合併)", ha='center', va='center', fontsize=8.5, fontweight='bold', color="#A80000")
    add_card(ax1, (50, 76), 45, 7, "#FDE7E9", "#E81123", radius=0.0)
    ax1.text(72.5, 79.5, "計費與貨況 (合併)", ha='center', va='center', fontsize=8.5, fontweight='bold', color="#A80000")
    
    sub_headers = ["提單號", "船名/航次", "裝卸港", "櫃號/尺寸", "運費(USD)", "狀態"]
    w_sub = [14, 18, 13, 17, 15, 13]
    cur_x = 5
    for sh, w in zip(sub_headers, w_sub):
        add_card(ax1, (cur_x, 70), w, 6, "#F3F2F1", "#A6A6A6", radius=0.0)
        ax1.text(cur_x+w/2, 73, sh, ha='center', va='center', fontsize=8, color="#323130")
        cur_x += w
        
    # Bad Data Rows
    bad_rows = [
        ["EMC-001", "EVER GIVEN\n0142-052W", "高雄/LAX", "EMCU123 (40HC)", "$2,450", "已付"],
        ["WHL-088", "WAN HAI 315", "基隆/上海", "WHLU987 (20GP)", "1200 USD", "到付"],
        ["【小計】", "共 2 筆提單", "--", "--", "$3,650", "--"], # Subtotal mixed in!
        ["", "", "", "", "", ""], # Blank line!
        ["YML-105", "YM WELLBEING", "高雄/鹿特丹", "YMLU456 (40GP)", "EUR 3,100", "已付"],
        ["OOCL-302", "OOCL POLAND", "高雄/長堤", "OOLU234 (40HC)", "$2,600", "已付"]
    ]
    
    y = 64
    for r_idx, r in enumerate(bad_rows):
        cur_x = 5
        is_subtotal = "小計" in r[0]
        is_blank = r[0] == ""
        bg = "#FFF0D4" if is_subtotal else ("#EAEAEA" if is_blank else "#FFFFFF")
        for val, w in zip(r, w_sub):
            add_card(ax1, (cur_x, y), w, 6, bg, "#D2D0CE", radius=0.0)
            fw = 'bold' if is_subtotal else 'normal'
            fc = "#D83B01" if is_subtotal else "#323130"
            ax1.text(cur_x+w/2, y+3, val, ha='center', va='center', fontsize=7.5, fontweight=fw, color=fc)
            cur_x += w
        y -= 6
        
    # Bad flaws labels
    ax1.text(50, 18, "[禁止] 嚴重致命缺陷清單：\n1. 頂部大標題跨欄置中，破壞欄位參照\n2. 雙層表頭合併儲存格，無法啟用篩選與樞紐分析\n3. 多項屬性擠在同一欄 (船名/航次、裝卸港口)\n4. 資料中插入「小計列」與「空白列」，中斷資料流\n5. 金額混雜幣別符號($/USD/EUR)，無法數值加總", 
             ha='center', va='center', fontsize=8.5, color="#A80000",
             bbox=dict(boxstyle="round,pad=0.5", fc="#FFF0F0", ec="#E81123", lw=1))
             
    # RIGHT: GOOD TABLE (第一正規化 / Tidy Data 標準二維資料表)
    add_card(ax2, (2, 2), 96, 96, "#F4FBF6", "#107C41", radius=0.4)
    ax2.text(50, 94, "[正確] 合格的資料表結構 (一維扁平化 - 標準資料庫結構)", ha='center', va='center', fontsize=11, fontweight='bold', color="#107C41")
    
    good_headers = ["提單號", "船名", "航次", "裝貨港", "卸貨港", "櫃號", "櫃型", "運費(USD)", "狀態"]
    w_good = [11, 13, 10, 8, 8, 12, 7, 11, 10]
    
    cur_x = 5
    for gh, w in zip(good_headers, w_good):
        add_card(ax2, (cur_x, 82), w, 7, "#107C41", "#0B552C", radius=0.0)
        ax2.text(cur_x+w/2, 85.5, gh, ha='center', va='center', fontsize=8, fontweight='bold', color="white")
        cur_x += w
        
    good_rows = [
        ["EMC-001", "EVER GIVEN", "0142-052W", "TWKHH", "USLAX", "EMCU1234567", "40HC", "2450.00", "PREPAID"],
        ["WHL-088", "WAN HAI 315", "0089-012N", "TWKEL", "CNSHA", "WHLU9876543", "20GP", "1200.00", "COLLECT"],
        ["YML-105", "YM WELLBEING", "0214-088E", "TWKHH", "NLRTM", "YMLU4567890", "40GP", "3100.00", "PREPAID"],
        ["OOCL-302", "OOCL POLAND", "0055-031W", "TWKHH", "USLGB", "OOLU2345678", "40HC", "2600.00", "PREPAID"],
        ["MSK-991", "MAERSK MC", "0411-002W", "TWKEL", "DEHAM", "MSKU3456789", "40HC", "3400.00", "PREPAID"],
        ["CMA-412", "CMA CGM ANTO", "0918-077E", "TWKHH", "FRLEH", "CMAU8765432", "20GP", "1450.00", "COLLECT"]
    ]
    
    y = 75
    for r_idx, r in enumerate(good_rows):
        cur_x = 5
        bg = "#F2F5F8" if r_idx % 2 == 1 else "#FFFFFF"
        for c_idx, (val, w) in enumerate(zip(r, w_good)):
            add_card(ax2, (cur_x, y), w, 6, bg, "#D2D0CE", radius=0.0)
            align = 'right' if c_idx == 7 else 'center'
            t_x = cur_x + w - 1 if align=='right' else cur_x + w/2
            ax2.text(t_x, y+3, val, ha=align, va='center', fontsize=7.5, color="#201F1E")
            cur_x += w
        y -= 6
        
    ax2.text(50, 18, "[優勢] 標準一維資料結構優勢：\n1. 單一列標題(Header)，無任何合併儲存格\n2. 一列代表一筆獨立的航運託運紀錄 (Record)\n3. 一欄只存放單一原子屬性 (船名與航次徹底分離)\n4. 資料連續無空白列、無中途小計，利於樞紐分析\n5. 金額為純數值型態，幣別以欄位獨立標註", 
             ha='center', va='center', fontsize=8.5, color="#0B552C",
             bbox=dict(boxstyle="round,pad=0.5", fc="#E6F2EA", ec="#107C41", lw=1))
             
    plt.tight_layout()
    plt.savefig("images/fig2_data_hygiene_good_vs_bad.png", dpi=200, bbox_inches='tight')
    plt.close()
    print("Generated fig2_data_hygiene_good_vs_bad.png")

# ========================================================
# Figure 3: Flash Fill (Ctrl+E) Demo
# ========================================================
def create_fig3():
    fig, ax = plt.subplots(figsize=(10, 5.5), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    add_card(ax, (2, 2), 96, 96, "#FFFFFF", "#1F497D", radius=0.3)
    ax.text(50, 93, "Excel 智慧特技：快速填入 (Flash Fill) 快捷鍵【Ctrl + E】操作拆解", ha='center', va='center', fontsize=11, fontweight='bold', color="#1F497D")
    
    # Spreadsheet Simulation
    cols = ["", "C 欄 (原始船名航次)", "D 欄 (目標船名 Vessel)", "E 欄 (目標航次 Voyage)"]
    widths = [6, 36, 26, 24]
    
    cur_x = 4
    for c_name, w in zip(cols, widths):
        add_card(ax, (cur_x, 78), w, 6, "#1F497D" if c_name else "#A6A6A6", "#102F54", radius=0.0)
        ax.text(cur_x+w/2, 81, c_name, ha='center', va='center', fontsize=8.5, fontweight='bold', color="white")
        cur_x += w
        
    data_demo = [
        ["1", "EVER GIVEN / 0142-052W", "EVER GIVEN", "0142-052W", "手動輸入第 1 筆範本 (人工打字)"],
        ["2", "WAN HAI 315 / 0089-012N", "WAN HAI 315", "0089-012N", "游標停在 D3 按下【Ctrl + E】即刻秒速填滿！"],
        ["3", "YM WELLBEING / 0214-088E", "YM WELLBEING", "0214-088E", "Excel 自動嗅探規律：擷取斜線左側為船名"],
        ["4", "OOCL POLAND / 0055-031W", "OOCL POLAND", "0055-031W", "Excel 自動嗅探規律：擷取斜線右側為航次"],
        ["5", "MAERSK MC-KINNEY / 0411-002W", "MAERSK MC-KINNEY", "0411-002W", "即使船名包含連字號符號也能精準辨識！"]
    ]
    
    y = 71
    for r_idx, r in enumerate(data_demo):
        cur_x = 4
        # highlight row 1 (manual input) vs rows 2-5 (auto flash fill)
        row_bg = "#FFF8E7" if r_idx == 0 else "#EBF3FB"
        for c_idx in range(4):
            w = widths[c_idx]
            val = r[c_idx]
            border_c = "#0078D4" if (r_idx > 0 and c_idx in [2, 3]) else "#D2D0CE"
            add_card(ax, (cur_x, y), w, 6.5, row_bg, border_c, radius=0.0)
            align = 'center' if c_idx == 0 else 'left'
            t_x = cur_x + w/2 if align=='center' else cur_x + 2
            fw = 'bold' if (r_idx == 0 and c_idx in [2, 3]) else 'normal'
            fc = "#D83B01" if (r_idx == 0 and c_idx in [2, 3]) else "#201F1E"
            ax.text(t_x, y+3.25, val, ha=align, va='center', fontsize=8, fontweight=fw, color=fc)
            cur_x += w
        y -= 6.5
        
    # Steps Box below
    add_card(ax, (4, 5), 92, 32, "#F8F9FA", "#0078D4", radius=0.3)
    ax.text(6, 32, "[提示] 航運實務新手操作 3 步驟口訣：", fontsize=9.5, fontweight='bold', color="#0078D4")
    ax.text(6, 25, "步驟 1【建立範本】：在緊鄰原始欄位的 D2 儲存格，親手輸入第一筆期望的乾淨結果『EVER GIVEN』。", fontsize=8.5, color="#323130")
    ax.text(6, 19, "步驟 2【快捷啟動】：移至下一格 D3，直接在鍵盤上按下萬能快捷鍵【Ctrl + E】（或點擊『資料』頁籤 ->『快速填入』）。", fontsize=8.5, color="#323130")
    ax.text(6, 13, "步驟 3【複檢查核】：Excel 即刻運用機器學習模型向下比對模式，自動為全欄 300 筆資料完整拆解！在 E 欄重複同樣動作拆航次。", fontsize=8.5, color="#323130")
    ax.text(6, 7, "[注意] 注意事項：快速填入產出為靜態文字；若原始資料事後變動，填入結果不會自動連動更新（需重按一次 Ctrl+E）。", fontsize=8, fontweight='bold', color="#D83B01")
    
    plt.tight_layout()
    plt.savefig("images/fig3_flash_fill_ctrl_e.png", dpi=200, bbox_inches='tight')
    plt.close()
    print("Generated fig3_flash_fill_ctrl_e.png")

# ========================================================
# Figure 4: Text to Columns Wizard (資料剖析精靈)
# ========================================================
def create_fig4():
    fig, ax = plt.subplots(figsize=(11, 6), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    add_card(ax, (2, 2), 96, 96, "#FFFFFF", "#107C41", radius=0.3)
    ax.text(50, 94, "功能選項示範：【資料剖析】精靈 3 步驟實戰分解 (Text to Columns)", ha='center', va='center', fontsize=11, fontweight='bold', color="#107C41")
    
    # 3 Step Panels
    step_boxes = [
        ((4, 10), 28, 78, "步驟 1 / 3：選擇原始資料類型", "#F2F5F8", [
            "● 分隔符號 (Delimited)",
            "  (例如逗號或定位字元分隔)",
            "○ 固定寬度 (Fixed Width)",
            "",
            "實務提示：",
            "船名航次以斜線『/』隔開，",
            "故此處請務必點選【分隔符號】。"
        ]),
        ((36, 10), 28, 78, "步驟 2 / 3：設定分隔符號", "#F2F5F8", [
            "核取方塊選項：",
            "[  ] 定位字元    [  ] 分號",
            "[  ] 逗號        [  ] 空格",
            "[V] 其他(O):  [ / ]",
            "",
            "資料預覽視窗：",
            "┌────────────┬──────────┐",
            "│ EVER GIVEN │0142-052W │",
            "│ WAN HAI 315│0089-012N │",
            "└────────────┴──────────┘"
        ]),
        ((68, 10), 28, 78, "步驟 3 / 3：欄位資料格式與目的端", "#F2F5F8", [
            "欄面的資料格式：",
            "● 一般(G)    ○ 文字(T)",
            "○ 日期(D):   [ YMD ▼ ]",
            "○ 不匯入此欄 (略過)",
            "",
            "目的端(Destination)：",
            "設定為: [$D$2]",
            "※保留原始 C 欄不受覆蓋破壞！"
        ])
    ]
    
    for (xy, w, h, title, bg, lines) in step_boxes:
        add_card(ax, xy, w, h, bg, "#0078D4", radius=0.2)
        # title bar
        add_card(ax, (xy[0], xy[1]+h-7), w, 7, "#0078D4", "#004578", radius=0.1)
        ax.text(xy[0]+w/2, xy[1]+h-3.5, title, ha='center', va='center', fontsize=8.5, fontweight='bold', color="white")
        
        # text lines
        cur_y = xy[1] + h - 11
        for line in lines:
            fw = 'bold' if ("●" in line or "[V]" in line or "目的端" in line or "提示" in line) else 'normal'
            fc = "#004578" if "●" in line or "[V]" in line else ("#D83B01" if "目的端" in line or "保留" in line else "#323130")
            ax.text(xy[0]+2, cur_y, line, va='center', fontsize=7.5, fontweight=fw, color=fc, fontfamily='Microsoft JhengHei' if "│" in line else 'Microsoft JhengHei')
            cur_y -= 4.2
            
    # Arrow between Step 1 and Step 2
    ax.annotate("", xy=(35, 50), xytext=(32.5, 50),
                arrowprops=dict(facecolor='#107C41', edgecolor='#0B552C', width=2, headwidth=7))
    # Arrow between Step 2 and Step 3
    ax.annotate("", xy=(67, 50), xytext=(64.5, 50),
                arrowprops=dict(facecolor='#107C41', edgecolor='#0B552C', width=2, headwidth=7))
                
    # Footer Action
    add_card(ax, (4, 3), 92, 5.5, "#FFF4CE", "#D83B01", radius=0.1)
    ax.text(50, 5.75, "[目標] 專家經驗談：剖析日期欄位時，步驟 3 可直接點選【日期】並指定【YMD】，Excel 會自動將文字轉換為實體日期序列值！", ha='center', va='center', fontsize=8, fontweight='bold', color="#7A0000")
    
    plt.tight_layout()
    plt.savefig("images/fig4_text_to_columns_wizard.png", dpi=200, bbox_inches='tight')
    plt.close()
    print("Generated fig4_text_to_columns_wizard.png")

# ========================================================
# Figure 5: Remove Duplicates Dialog (移除重複項)
# ========================================================
def create_fig5():
    fig, ax = plt.subplots(figsize=(9, 5.5), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    # Background Canvas
    add_card(ax, (2, 2), 96, 96, "#F3F2F1", "#A6A6A6", radius=0.3)
    
    # Dialog Window
    add_card(ax, (15, 12), 70, 78, "#FFFFFF", "#605E5C", radius=0.2)
    # Dialog Title Bar
    add_card(ax, (15, 83), 70, 7, "#FFFFFF", "#D2D0CE", radius=0.1)
    ax.text(18, 86.5, "移除重複項 (Remove Duplicates)", va='center', fontsize=9.5, fontweight='bold', color="#201F1E")
    ax.text(82, 86.5, "X", va='center', fontsize=11, color="#605E5C")
    
    # Explanation
    ax.text(18, 78, "若要刪除重複值，請選取一或多個包含重複項目的欄位。", fontsize=8, color="#605E5C")
    
    # Header checkbox
    add_card(ax, (18, 70), 3, 3, "#0078D4", "#004578", radius=0.0)
    ax.text(19.5, 71.5, "V", ha='center', va='center', fontsize=7, color="white", fontweight='bold')
    ax.text(23, 71.5, "我的資料有標頭(H)", va='center', fontsize=8.5, fontweight='bold', color="#201F1E")
    
    # Select / Deselect Buttons
    add_card(ax, (52, 69.5), 15, 4, "#F3F2F1", "#8A8886", radius=0.1)
    ax.text(59.5, 71.5, "全選(S)", ha='center', va='center', fontsize=7.5, color="#201F1E")
    add_card(ax, (68, 69.5), 15, 4, "#F3F2F1", "#8A8886", radius=0.1)
    ax.text(75.5, 71.5, "取消全選(U)", ha='center', va='center', fontsize=7.5, color="#201F1E")
    
    # Columns list box
    add_card(ax, (18, 25), 64, 42, "#FFFFFF", "#8A8886", radius=0.0)
    col_items = [
        ("[ ]", "項次 (序號)", False),
        ("[V]", "提單號碼 (B/L No.)  ← 【核心主鍵！請單獨勾選此項】", True),
        ("[ ]", "船名航次 (Vessel/Voyage)", False),
        ("[ ]", "裝貨港 (POL)", False),
        ("[ ]", "貨櫃號碼 (Container No.)", False),
        ("[ ]", "海運費 (Ocean Freight)", False)
    ]
    cur_y = 62
    for box, label, is_key in col_items:
        bg_bar = "#C7E0F4" if is_key else "#FFFFFF"
        if is_key:
            add_card(ax, (19, cur_y-1.5), 62, 5, bg_bar, "#0078D4", radius=0.1)
        ax.text(21, cur_y+1, box, va='center', fontsize=9, fontweight='bold', color="#0078D4" if is_key else "#605E5C")
        ax.text(25, cur_y+1, label, va='center', fontsize=8, fontweight='bold' if is_key else 'normal', color="#004578" if is_key else "#323130")
        cur_y -= 6.5
        
    # Dialog Bottom Buttons
    add_card(ax, (52, 15), 14, 5, "#0078D4", "#004578", radius=0.1)
    ax.text(59, 17.5, "確定 (OK)", ha='center', va='center', fontsize=8.5, fontweight='bold', color="white")
    add_card(ax, (68, 15), 14, 5, "#F3F2F1", "#8A8886", radius=0.1)
    ax.text(75, 17.5, "取消", ha='center', va='center', fontsize=8.5, color="#201F1E")
    
    # Result Prompt Box
    add_card(ax, (20, 3), 60, 8, "#E6F2EA", "#107C41", radius=0.2)
    ax.text(50, 7, "系統提示：已找到並移除 1 個重複的值 (Row 12 提單 MSK-2026-991)；\n保留了 29 個唯一值。", ha='center', va='center', fontsize=7.5, fontweight='bold', color="#0B552C")
    
    plt.tight_layout()
    plt.savefig("images/fig5_remove_duplicates_dialog.png", dpi=200, bbox_inches='tight')
    plt.close()
    print("Generated fig5_remove_duplicates_dialog.png")

# ========================================================
# Figure 6: Find and Replace Dialog (尋找與取代 Ctrl+H)
# ========================================================
def create_fig6():
    fig, ax = plt.subplots(figsize=(9, 5), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    add_card(ax, (2, 2), 96, 96, "#F3F2F1", "#A6A6A6", radius=0.3)
    
    # Dialog Box
    add_card(ax, (12, 10), 76, 80, "#FFFFFF", "#605E5C", radius=0.2)
    # Title bar
    add_card(ax, (12, 83), 76, 7, "#FFFFFF", "#D2D0CE", radius=0.1)
    ax.text(15, 86.5, "尋找與取代 (Find and Replace)  - 【Ctrl + H】", va='center', fontsize=9.5, fontweight='bold', color="#201F1E")
    
    # Tabs
    add_card(ax, (16, 75), 14, 5, "#F3F2F1", "#8A8886", radius=0.0)
    ax.text(23, 77.5, "尋找(D)", ha='center', va='center', fontsize=8, color="#605E5C")
    add_card(ax, (30, 75), 14, 5, "#FFFFFF", "#0078D4", radius=0.0)
    ax.text(37, 77.5, "取代(P)", ha='center', va='center', fontsize=8, fontweight='bold', color="#0078D4")
    
    # Inputs
    ax.text(16, 66, "尋找目標(N):", va='center', fontsize=8.5, color="#201F1E")
    add_card(ax, (32, 63), 42, 6, "#FFFFFF", "#0078D4", radius=0.0)
    ax.text(34, 66, " kg", va='center', fontsize=9, fontweight='bold', color="#D83B01", fontfamily='Microsoft JhengHei')
    ax.text(75, 66, "(示範清除單位或空白)", va='center', fontsize=7.5, color="#605E5C")
    
    ax.text(16, 54, "取代為(E):", va='center', fontsize=8.5, color="#201F1E")
    add_card(ax, (32, 51), 42, 6, "#FFFFFF", "#8A8886", radius=0.0)
    ax.text(34, 54, "【完全留空】", va='center', fontsize=8.5, color="#107C41")
    ax.text(75, 54, "(替換為空即為刪除)", va='center', fontsize=7.5, color="#605E5C")
    
    # Options & Buttons
    add_card(ax, (16, 38), 18, 5, "#F3F2F1", "#8A8886", radius=0.1)
    ax.text(25, 40.5, "全部取代(A)", ha='center', va='center', fontsize=8, fontweight='bold', color="#004578")
    add_card(ax, (36, 38), 14, 5, "#F3F2F1", "#8A8886", radius=0.1)
    ax.text(43, 40.5, "取代(R)", ha='center', va='center', fontsize=8, color="#201F1E")
    add_card(ax, (52, 38), 16, 5, "#F3F2F1", "#8A8886", radius=0.1)
    ax.text(60, 40.5, "尋找全部(I)", ha='center', va='center', fontsize=8, color="#201F1E")
    add_card(ax, (70, 38), 16, 5, "#F3F2F1", "#8A8886", radius=0.1)
    ax.text(78, 40.5, "關閉", ha='center', va='center', fontsize=8, color="#201F1E")
    
    # Practical scenarios bullet list
    add_card(ax, (16, 14), 68, 20, "#FFF4CE", "#D83B01", radius=0.2)
    ax.text(18, 30, "[重點] 航運行政 3 大高頻『全部取代』實用情境：", fontsize=8, fontweight='bold', color="#7A0000")
    ax.text(18, 24, "1. 剔除重量字串尾綴：尋找目標輸入『kg』或『KGS』，取代留空。", fontsize=7.5, color="#323130")
    ax.text(18, 19, "2. 剔除運費幣別代碼：尋找目標輸入『$』、『USD』或『EUR』，取代留空。", fontsize=7.5, color="#323130")
    ax.text(18, 14, "3. 消除字串內部空格：尋找目標按一下【空白鍵】，取代留空，批次乾淨拔除空白！", fontsize=7.5, color="#323130")
    
    plt.tight_layout()
    plt.savefig("images/fig6_find_and_replace_dialog.png", dpi=200, bbox_inches='tight')
    plt.close()
    print("Generated fig6_find_and_replace_dialog.png")

# ========================================================
# Figure 7: Data Validation Dialog (資料驗證防錯)
# ========================================================
def create_fig7():
    fig, ax = plt.subplots(figsize=(9, 5.5), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    add_card(ax, (2, 2), 96, 96, "#F3F2F1", "#A6A6A6", radius=0.3)
    
    # Dialog
    add_card(ax, (12, 10), 76, 82, "#FFFFFF", "#605E5C", radius=0.2)
    # Title
    add_card(ax, (12, 85), 76, 7, "#FFFFFF", "#D2D0CE", radius=0.1)
    ax.text(15, 88.5, "資料驗證 (Data Validation) - 防止輸入髒資料的源頭防線", va='center', fontsize=9.5, fontweight='bold', color="#201F1E")
    
    # Tabs
    add_card(ax, (16, 77), 16, 5, "#FFFFFF", "#0078D4", radius=0.0)
    ax.text(24, 79.5, "設定(S)", ha='center', va='center', fontsize=8, fontweight='bold', color="#0078D4")
    add_card(ax, (33, 77), 16, 5, "#F3F2F1", "#8A8886", radius=0.0)
    ax.text(41, 79.5, "提示訊息(I)", ha='center', va='center', fontsize=8, color="#605E5C")
    add_card(ax, (50, 77), 16, 5, "#F3F2F1", "#8A8886", radius=0.0)
    ax.text(58, 79.5, "錯誤警訊(E)", ha='center', va='center', fontsize=8, color="#605E5C")
    
    # Settings Area
    ax.text(16, 70, "驗證準則：", fontsize=8.5, fontweight='bold', color="#201F1E")
    ax.text(16, 63, "儲存格允許(A):", va='center', fontsize=8, color="#201F1E")
    add_card(ax, (35, 60), 38, 6, "#FFFFFF", "#0078D4", radius=0.0)
    ax.text(37, 63, "清單 (List)                       ▼", va='center', fontsize=8.5, fontweight='bold', color="#004578")
    
    ax.text(16, 50, "來源(S):", va='center', fontsize=8, color="#201F1E")
    add_card(ax, (35, 47), 48, 6, "#FFFFFF", "#0078D4", radius=0.0)
    ax.text(37, 50, "PREPAID, COLLECT", va='center', fontsize=9, fontweight='bold', color="#D83B01", fontfamily='Microsoft JhengHei')
    ax.text(16, 41, "(或選取港口代碼參照區: =Port_Master!$A$2:$A$18)", fontsize=7.5, color="#605E5C")
    
    # Dropdown UI Simulation
    add_card(ax, (16, 16), 70, 22, "#EBF3FB", "#0078D4", radius=0.2)
    ax.text(18, 33, "[清單] 試算表端實體效果（限制行政人員只能從下拉清單點選）：", fontsize=8, fontweight='bold', color="#004578")
    
    add_card(ax, (20, 21), 28, 7, "#FFFFFF", "#0078D4", radius=0.1)
    ax.text(22, 24.5, "PREPAID        ▼", va='center', fontsize=9, fontweight='bold', color="#107C41")
    
    # Dropdown box opened
    add_card(ax, (52, 17), 30, 14, "#FFFFFF", "#8A8886", radius=0.0)
    ax.text(54, 27, "PREPAID (預付運費)", va='center', fontsize=8, color="#201F1E")
    ax.text(54, 21, "COLLECT (到付運費)", va='center', fontsize=8, color="#201F1E")
    
    plt.tight_layout()
    plt.savefig("images/fig7_data_validation_dialog.png", dpi=200, bbox_inches='tight')
    plt.close()
    print("Generated fig7_data_validation_dialog.png")

# ========================================================
# Figure 8: Filter and Sort UI (自動篩選與排序)
# ========================================================
def create_fig8():
    fig, ax = plt.subplots(figsize=(10, 5.5), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    add_card(ax, (2, 2), 96, 96, "#FFFFFF", "#1F497D", radius=0.3)
    ax.text(50, 93, "資料品質快速檢驗：【自動篩選】(Ctrl+Shift+L) 與下拉清單排錯診斷", ha='center', va='center', fontsize=11, fontweight='bold', color="#1F497D")
    
    # Table Header with Filter Arrows
    headers = ["提單號碼", "裝貨港 ▽", "貨櫃號碼", "海運費(USD) ▽", "付款狀態 ▽"]
    w_hd = [20, 18, 20, 20, 18]
    cur_x = 4
    for h, w in zip(headers, w_hd):
        add_card(ax, (cur_x, 78), w, 7, "#1F497D", "#102F54", radius=0.0)
        ax.text(cur_x+w/2, 81.5, h, ha='center', va='center', fontsize=8.5, fontweight='bold', color="white")
        cur_x += w
        
    # Dropdown Menu Simulation for "裝貨港 ▽"
    add_card(ax, (24, 15), 45, 62, "#FFFFFF", "#605E5C", radius=0.1)
    # Menu Items
    ax.text(26, 72, "↑ 從 A 到 Z 排序", va='center', fontsize=8, color="#201F1E")
    ax.text(26, 66, "↓ 從 Z 到 A 排序", va='center', fontsize=8, color="#201F1E")
    ax.text(26, 60, "依色彩篩選           ▶", va='center', fontsize=8, color="#201F1E")
    ax.text(26, 54, "文字篩選             ▶", va='center', fontsize=8, color="#201F1E")
    
    # Search Box inside filter
    add_card(ax, (26, 45), 41, 5, "#F3F2F1", "#8A8886", radius=0.1)
    ax.text(28, 47.5, " 搜尋...", va='center', fontsize=7.5, color="#605E5C")
    
    # Checkbox items exposing dirty data flaws!
    flaw_items = [
        ("[V] (全選)", "#201F1E"),
        ("[V] TWKEL", "#107C41"),
        ("[V] TWKHH", "#107C41"),
        ("[V]  TWKHH   (⚠️抓包！前後有多餘空格)", "#D83B01"),
        ("[V] twkhh    (⚠️抓包！全小寫未標準化)", "#D83B01"),
        ("[V] (空格)    (⚠️抓包！遺漏值未填)", "#E81123")
    ]
    cur_y = 39
    for item, col in flaw_items:
        fw = 'bold' if "⚠️" in item else 'normal'
        ax.text(27, cur_y, item, va='center', fontsize=7.5, fontweight=fw, color=col)
        cur_y -= 4.5
        
    # Callout pointing to dirty values in filter
    ax.annotate("關鍵技巧：\n點擊篩選箭頭展開清單，\n所有大小寫不一、夾帶空格、\n空白遺漏值立刻無所遁形！",
                xy=(55, 30), xytext=(78, 40),
                arrowprops=dict(facecolor='#D83B01', edgecolor='#A80000', width=1.5, headwidth=6),
                bbox=dict(boxstyle="round,pad=0.4", fc="#FFF4CE", ec="#D83B01", lw=1),
                fontsize=8, fontweight='bold', color="#7A0000", ha='center', va='center')
                
    plt.tight_layout()
    plt.savefig("images/fig8_filter_and_sort_ui.png", dpi=200, bbox_inches='tight')
    plt.close()
    print("Generated fig8_filter_and_sort_ui.png")

# ========================================================
# Figure 9: Text Formulas in Action (文字清洗函數圖解)
# ========================================================
def create_fig9():
    fig, ax = plt.subplots(figsize=(11, 5.5), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    add_card(ax, (2, 2), 96, 96, "#FFFFFF", "#1F497D", radius=0.3)
    ax.text(50, 93, "核心公式圖解：文字清洗五大天王函數實戰鏈 (TRIM / UPPER / PROPER / SUBSTITUTE / MID)", ha='center', va='center', fontsize=10.5, fontweight='bold', color="#1F497D")
    
    cards = [
        ((4, 52), 43, 34, "=TRIM(文字)", "清除字串首尾空格", [
            "輸入：'  EMC-2026-001  ' (前後有贅字空格)",
            "輸出：'EMC-2026-001'",
            "特性：自動移除首尾全部空白，",
            "      字詞中間若有多個連續空格僅保留單一格。"
        ], "#0078D4"),
        ((53, 52), 43, 34, "=UPPER(文字) / =LOWER / =PROPER", "大小寫格式標準化", [
            "輸入：'twkhh' 或 'cnsHa'",
            "公式：=UPPER(TRIM(D2))",
            "輸出：'TWKHH' / 'CNSHA'",
            "特性：確保港口 UN/LOCODE 與貨櫃代碼完全符合",
            "      國際大寫規範，消除電腦比對盲區。"
        ], "#107C41"),
        ((4, 10), 43, 36, "=SUBSTITUTE(文字, 舊字, 新字)", "精準替換或拔除指定字元", [
            "輸入：'18,500 kg' (帶單位無法計算)",
            "公式：=SUBSTITUTE(H2, \"kg\", \"\")",
            "進階：雙重嵌套消除單位與逗號",
            "      =VALUE(SUBSTITUTE(H2, \"kg\", \"\"))",
            "輸出：實體數值 18500"
        ], "#D83B01"),
        ((53, 10), 43, 36, "=LEFT / =MID / =RIGHT / =FIND", "動態座標截取字串", [
            "輸入：'EVER GIVEN / 0142-052W'",
            "船名：=TRIM(LEFT(C2, FIND(\"/\", C2)-1))",
            "航次：=TRIM(MID(C2, FIND(\"/\", C2)+1, 50))",
            "原理：FIND 自動定位斜線『/』位置，",
            "      LEFT/MID 依據動態切點精確拆解欄位！"
        ], "#5C2D91")
    ]
    
    for (xy, w, h, title, subtitle, lines, color) in cards:
        add_card(ax, xy, w, h, "#F8F9FA", color, radius=0.2)
        add_card(ax, (xy[0], xy[1]+h-7), w, 7, color, color, radius=0.1)
        ax.text(xy[0]+2, xy[1]+h-3.5, title, va='center', fontsize=8.5, fontweight='bold', color="white", fontfamily='Microsoft JhengHei')
        ax.text(xy[0]+w-2, xy[1]+h-3.5, subtitle, ha='right', va='center', fontsize=7.5, color="white")
        
        cur_y = xy[1] + h - 11
        for l in lines:
            fw = 'bold' if ("輸出" in l or "公式" in l) else 'normal'
            fc = color if ("輸出" in l or "公式" in l) else "#323130"
            ax.text(xy[0]+2, cur_y, l, va='center', fontsize=7.5, fontweight=fw, color=fc)
            cur_y -= 5
            
    plt.tight_layout()
    plt.savefig("images/fig9_text_formulas_workflow.png", dpi=200, bbox_inches='tight')
    plt.close()
    print("Generated fig9_text_formulas_workflow.png")

# ========================================================
# Figure 10: Date Cleaning & Transit Calculation
# ========================================================
def create_fig10():
    fig, ax = plt.subplots(figsize=(10, 5), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    add_card(ax, (2, 2), 96, 96, "#FFFFFF", "#1F497D", radius=0.3)
    ax.text(50, 93, "日期前處理與航運時效運算：【文字日期 -> 標準日期序列 -> 航程天數】", ha='center', va='center', fontsize=10.5, fontweight='bold', color="#1F497D")
    
    # 3-Stage Pipeline
    stages = [
        ((4, 25), 26, 58, "Stage 1: 混亂文字日期", "#FFF0F0", "#E81123", [
            "航運常見混亂格式：",
            "• '2026/10/05'",
            "• '2026-10-22'",
            "• '10-10-2026'",
            "• '2026.11.08'",
            "",
            "[注意] 致命痛點：",
            "靠左對齊之文字型態",
            "無法直接相減計算天數！"
        ]),
        ((37, 25), 26, 58, "Stage 2: 轉換為日期序列值", "#F4FBF6", "#107C41", [
            "清洗轉換神器：",
            "1. 資料剖析 (指定YMD)",
            "2. DATEVALUE 函數：",
            "   =DATEVALUE(",
            "    SUBSTITUTE(",
            "    SUBSTITUTE(J2,",
            "    \".\",\"/\"),\"-\",\"/\"))",
            "",
            "底層本質：Excel 以",
            "1900/1/1 起算之天數儲存"
        ]),
        ((70, 25), 26, 58, "Stage 3: 航運業務天數計算", "#EBF3FB", "#0078D4", [
            "標準航程天數公式：",
            "【航程】= ETA - ETD",
            "",
            "實務到港滯留天數：",
            "【滯留】= GateOut - ATA",
            "",
            "超期罰款天數預警：",
            "=MAX(0, 滯留 - 免租期)",
            "直接得出滯港費判定！"
        ])
    ]
    
    for (xy, w, h, title, bg, border, lines) in stages:
        add_card(ax, xy, w, h, bg, border, radius=0.2)
        add_card(ax, (xy[0], xy[1]+h-7), w, 7, border, border, radius=0.1)
        ax.text(xy[0]+w/2, xy[1]+h-3.5, title, ha='center', va='center', fontsize=8.5, fontweight='bold', color="white")
        
        cur_y = xy[1] + h - 12
        for l in lines:
            fw = 'bold' if ("【" in l or "Stage" in l or "1." in l or "2." in l) else 'normal'
            fc = border if "【" in l else ("#D83B01" if "⚠️" in l else "#323130")
            ax.text(xy[0]+2, cur_y, l, va='center', fontsize=7.5, fontweight=fw, color=fc, fontfamily='Microsoft JhengHei' if "=" in l else 'Microsoft JhengHei')
            cur_y -= 4.8
            
    # Connecting Arrows
    ax.annotate("", xy=(36, 54), xytext=(31, 54),
                arrowprops=dict(facecolor='#107C41', edgecolor='#0B552C', width=2.5, headwidth=8))
    ax.annotate("", xy=(69, 54), xytext=(64, 54),
                arrowprops=dict(facecolor='#0078D4', edgecolor='#004578', width=2.5, headwidth=8))
                
    # Bottom Note
    add_card(ax, (4, 6), 92, 14, "#F8F9FA", "#605E5C", radius=0.1)
    ax.text(6, 15, "[提示] 航運初學者避坑心法：", fontsize=8, fontweight='bold', color="#004578")
    ax.text(6, 10, "當你看到『2026/10/15』靠儲存格左側時，代表它是偽裝成日期的『字串』；唯有靠右對齊且能參與加減運算的，才是真正的『日期序列值』！", fontsize=7.5, color="#323130")
    
    plt.tight_layout()
    plt.savefig("images/fig10_date_functions_and_transit.png", dpi=200, bbox_inches='tight')
    plt.close()
    print("Generated fig10_date_functions_and_transit.png")

# ========================================================
# Figure 11: XLOOKUP / VLOOKUP Port Code Mapping
# ========================================================
def create_fig11():
    fig, ax = plt.subplots(figsize=(10.5, 5.5), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    add_card(ax, (2, 2), 96, 96, "#FFFFFF", "#1F497D", radius=0.3)
    ax.text(50, 93, "現代動態查表關聯：【XLOOKUP / VLOOKUP】跨表串接港口基本檔 (Port Master)", ha='center', va='center', fontsize=10.5, fontweight='bold', color="#1F497D")
    
    # Left: Main Shipping Transaction Table
    add_card(ax, (4, 20), 42, 65, "#F2F5F8", "#0078D4", radius=0.2)
    ax.text(25, 81, "航運業務主表 (Transaction Table)", ha='center', va='center', fontsize=8.5, fontweight='bold', color="#004578")
    
    m_headers = ["提單號", "起運港(POL)", "起運港中文名稱(Lookup)"]
    m_w = [11, 13, 16]
    cur_x = 5
    for h, w in zip(m_headers, m_w):
        add_card(ax, (cur_x, 70), w, 6, "#0078D4", "#004578", radius=0.0)
        ax.text(cur_x+w/2, 73, h, ha='center', va='center', fontsize=7, fontweight='bold', color="white")
        cur_x += w
        
    m_rows = [
        ["EMC-001", "TWKHH", "高雄港 (自動帶出)"],
        ["WHL-088", "TWKEL", "基隆港 (自動帶出)"],
        ["YML-105", "TWKHH", "高雄港 (自動帶出)"],
        ["ONE-550", "TWTXG", "台中港 (自動帶出)"]
    ]
    y = 64
    for r_idx, r in enumerate(m_rows):
        cur_x = 5
        bg = "#FFF8E7" if r_idx == 0 else "#FFFFFF"
        for c_idx, (val, w) in enumerate(zip(r, m_w)):
            add_card(ax, (cur_x, y), w, 6, bg, "#D2D0CE", radius=0.0)
            fc = "#D83B01" if (r_idx==0 and c_idx==1) else ("#107C41" if (r_idx==0 and c_idx==2) else "#201F1E")
            fw = 'bold' if r_idx==0 else 'normal'
            ax.text(cur_x+w/2, y+3, val, ha='center', va='center', fontsize=7, fontweight=fw, color=fc)
            cur_x += w
        y -= 6
        
    # Right: Port Master Lookup Table
    add_card(ax, (54, 20), 42, 65, "#F4FBF6", "#107C41", radius=0.2)
    ax.text(75, 81, "港口代碼參照表 (Port_Master 工作表)", ha='center', va='center', fontsize=8.5, fontweight='bold', color="#0B552C")
    
    p_headers = ["港口代碼(LOCODE)", "港口中文名", "所屬國家/地區"]
    p_w = [14, 13, 13]
    cur_x = 55
    for h, w in zip(p_headers, p_w):
        add_card(ax, (cur_x, 70), w, 6, "#107C41", "#0B552C", radius=0.0)
        ax.text(cur_x+w/2, 73, h, ha='center', va='center', fontsize=7, fontweight='bold', color="white")
        cur_x += w
        
    p_rows = [
        ["TWKHH", "高雄港", "台灣 (Taiwan)"],
        ["TWKEL", "基隆港", "台灣 (Taiwan)"],
        ["TWTXG", "台中港", "台灣 (Taiwan)"],
        ["USLAX", "洛杉磯港", "美國 (USA)"],
        ["NLRTM", "鹿特丹港", "荷蘭 (Netherlands)"]
    ]
    y = 64
    for r_idx, r in enumerate(p_rows):
        cur_x = 55
        bg = "#E6F2EA" if r_idx == 0 else "#FFFFFF"
        for c_idx, (val, w) in enumerate(zip(r, p_w)):
            add_card(ax, (cur_x, y), w, 6, bg, "#D2D0CE", radius=0.0)
            fc = "#0B552C" if r_idx==0 else "#201F1E"
            fw = 'bold' if r_idx==0 else 'normal'
            ax.text(cur_x+w/2, y+3, val, ha='center', va='center', fontsize=7, fontweight=fw, color=fc)
            cur_x += w
        y -= 6
        
    # Curved arrow showing lookup connection
    con = ConnectionPatch(xyA=(23, 67), xyB=(62, 67), coordsA="data", coordsB="data",
                          axesA=ax, axesB=ax, arrowstyle="<->", shrinkA=5, shrinkB=5,
                          mutation_scale=15, fc="#D83B01", ec="#D83B01", lw=2)
    ax.add_artist(con)
    
    # Formula Display Box at Bottom
    add_card(ax, (4, 4), 92, 13, "#FFF4CE", "#D83B01", radius=0.2)
    ax.text(6, 12.5, "[提示] 新世代首選公式：=XLOOKUP(B2, Port_Master!$A$2:$A$18, Port_Master!$B$2:$B$18, \"未知港口\")", fontsize=8, fontweight='bold', color="#7A0000", fontfamily='Microsoft JhengHei')
    ax.text(6, 7.5, "[重點] 經典傳統相容公式：=VLOOKUP(B2, Port_Master!$A$2:$C$18, 2, FALSE)", fontsize=8, color="#323130", fontfamily='Microsoft JhengHei')
    
    plt.tight_layout()
    plt.savefig("images/fig11_xlookup_port_master.png", dpi=200, bbox_inches='tight')
    plt.close()
    print("Generated fig11_xlookup_port_master.png")

# ========================================================
# Figure 12: 300-Row Practice Master Architecture
# ========================================================
def create_fig12():
    fig, ax = plt.subplots(figsize=(11, 5.5), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    add_card(ax, (2, 2), 96, 96, "#FFFFFF", "#1F497D", radius=0.3)
    ax.text(50, 93, "航運大數據清洗專案架構地圖 (300 筆綜合實戰情境全流程管線)", ha='center', va='center', fontsize=11, fontweight='bold', color="#1F497D")
    
    pipeline_steps = [
        ((4, 15), 16, 72, "Step 1: 資料導入與結構審核", "#EBF3FB", "#0078D4", [
            "• 檢查一維表格規範",
            "• 凍結頂端列視窗",
            "• 啟用自動篩選",
            "• 排查全域空行空欄"
        ]),
        ((23, 15), 16, 72, "Step 2: 字串與代碼標準化", "#EBF3FB", "#0078D4", [
            "• B/L 與訂艙去首尾空格",
            "• 港口 LOCODE 全大寫",
            "• 貨櫃號碼 11 碼去空",
            "• 船名航次雙欄分離"
        ]),
        ((42, 15), 16, 72, "Step 3: 數值與幣別淨化", "#FFF4CE", "#D83B01", [
            "• 拔除毛重 'kg' 單位",
            "• 分離運費幣別與金額",
            "• 轉換文字型數字為數值",
            "• 統一千分位格式"
        ]),
        ((61, 15), 16, 72, "Step 4: 時效運算與碼表關聯", "#F4FBF6", "#107C41", [
            "• 統一多樣化日期格式",
            "• 計算航程天數(ETA-ETD)",
            "• 計算留港天數與滯港費",
            "• XLOOKUP 帶出港名/TEU"
        ]),
        ((80, 15), 16, 72, "Step 5: 重複排查與防錯稽核", "#F9F0FF", "#5C2D91", [
            "• 提單主鍵重複項剔除",
            "• 空白客戶補標記預設",
            "• 異常值條件格式警示",
            "• 產出標準分析主檔"
        ])
    ]
    
    for (xy, w, h, title, bg, border, items) in pipeline_steps:
        add_card(ax, xy, w, h, bg, border, radius=0.2)
        add_card(ax, (xy[0], xy[1]+h-9), w, 9, border, border, radius=0.1)
        ax.text(xy[0]+w/2, xy[1]+h-4.5, title, ha='center', va='center', fontsize=7.5, fontweight='bold', color="white")
        
        cur_y = xy[1] + h - 14
        for it in items:
            ax.text(xy[0]+1.5, cur_y, it, va='center', fontsize=7.5, color="#201F1E")
            cur_y -= 8
            
    # Connecting Arrows across pipeline
    for x in [20.5, 39.5, 58.5, 77.5]:
        ax.annotate("", xy=(x+2, 51), xytext=(x-0.5, 51),
                    arrowprops=dict(facecolor='#1F497D', edgecolor='#102F54', width=2, headwidth=6))
                    
    # Bottom Note
    add_card(ax, (4, 4), 92, 8, "#F3F2F1", "#8A8886", radius=0.1)
    ax.text(50, 8, "[目標] 核心宗旨：透過標準化前處理管線，將充滿雜訊的原始 EDI 營運紀錄，轉化為高可靠度之決策資料庫！", ha='center', va='center', fontsize=8, fontweight='bold', color="#1F497D")
    
    plt.tight_layout()
    plt.savefig("images/fig12_practice_300_workflow.png", dpi=200, bbox_inches='tight')
    plt.close()
    print("Generated fig12_practice_300_workflow.png")

if __name__ == "__main__":
    create_fig1()
    create_fig2()
    create_fig3()
    create_fig4()
    create_fig5()
    create_fig6()
    create_fig7()
    create_fig8()
    create_fig9()
    create_fig10()
    create_fig11()
    create_fig12()
    print("All 12 figures generated successfully in images/ directory.")
