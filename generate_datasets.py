import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import random
from datetime import datetime, timedelta

def style_header_row(ws, col_count, fill_color="1F497D", font_color="FFFFFF"):
    font = Font(name="Microsoft JhengHei", size=11, bold=True, color=font_color)
    fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='medium', color='1F497D'),
        bottom=Side(style='medium', color='1F497D')
    )
    for col in range(1, col_count + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = font
        cell.fill = fill
        cell.alignment = align
        cell.border = border
    ws.row_dimensions[1].height = 28

def style_data_rows(ws, row_count, col_count):
    font = Font(name="Microsoft JhengHei", size=10)
    border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    alt_fill = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    for row in range(2, row_count + 2):
        ws.row_dimensions[row].height = 20
        fill = alt_fill if row % 2 == 1 else white_fill
        for col in range(1, col_count + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = font
            cell.border = border
            if cell.fill.fill_type is None:
                cell.fill = fill

def autofit_columns(ws, max_len_cap=35):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            # approximate visual length for Chinese characters
            char_len = sum(2 if ord(c) > 127 else 1 for c in val_str)
            if char_len > max_len:
                max_len = char_len
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), max_len_cap)

# ==========================================
# 1. GENERATE 30-ROW DEMO DATASET
# ==========================================
def generate_30_demo():
    wb = openpyxl.Workbook()
    
    # ----------------- Sheet 1: Raw Dirty Data -----------------
    ws_raw = wb.active
    ws_raw.title = "航運原始資料_未清洗(Raw)"
    
    headers_raw = [
        "項次", "提單號碼(B/L No.)", "船名航次(Vessel/Voyage)", "裝貨港(POL)", "卸貨港(POD)", 
        "貨櫃號碼(Container No.)", "貨櫃尺寸", "貨物毛重(Gross Wt)", "海運費(Ocean Freight)", 
        "預計開航日(ETD)", "預計抵達日(ETA)", "運費付款狀態", "託運人名稱(Shipper)"
    ]
    ws_raw.append(headers_raw)
    
    # 30 rows designed with realistic dirty data flaws:
    # 1. Leading/trailing/irregular spaces
    # 2. Inconsistent casing (twkhh vs TWKHH vs twKhh)
    # 3. Delimited text in single column ("EVER GIVEN / 0124-045W")
    # 4. Text-formatted numbers / mixed currency strings ("$2,450.00", "USD 1800", "2200 USD", " 3,100 ")
    # 5. Inconsistent date formats ("2026/10/12", "2026-10-15", "12-10-2026", "2026.10.18", text dates)
    # 6. Duplicate rows (Row 5 and Row 12 duplicate B/L)
    # 7. Numbers stored as text (e.g. Gross Weight with 'kg' or string format)
    # 8. Missing values in Shipper or Payment status
    raw_data_30 = [
        [1, "  EMC-2026-001  ", "EVER GIVEN / 0142-052W", "twkhh", "USLAX", "EMCU1234567 ", "40HC", "18500 kg", "$2,450.00", "2026/10/05", "2026-10-22", "PAID", "宏達國際貿易股份有限公司"],
        [2, "WHL-2026-088", "WAN HAI 315 / 0089-012N", "TWKEL", "cnsHa", "WHLU9876543", "20GP", " 12,400 ", "USD 1200", "2026-10-06", "2026/10/10", "Collect", "台灣造紙工業股份有限公司"],
        [3, "YML-2026-105 ", "YM WELLBEING / 0214-088E", "twkhh", "NLRTM", "YMLU4567890", "40GP", "21300", "EUR 3,100", "10-10-2026", "2026.11.08", "paid", "長興材料工業股份有限公司"],
        [4, "OOCL-2026-302", "OOCL POLAND / 0055-031W", "TWKHH", "USLGB", "OOLU2345678", "40HC", "19,800kg", "$2,600", "2026/10/11", "2026-10-28", "PREPAID", "和大工業股份有限公司"],
        [5, "MSK-2026-991", "MAERSK MC-KINNEY / 0411-002W", "TWKEL", "DEHAM", "MSKU3456789", "40HC", "22500", "3400 USD", "2026.10.12", "2026/11/15", "PAID", "日月光半導體製造股份有限公司"],
        [6, "CMA-2026-412", "CMA CGM ANTOINE / 0918-077E", "twkhh", "FRLEH", "CMAU8765432", "20GP", "9800", "$1,450", "2026/10/13", "2026-11-18", "Collect", "台灣塑膠工業股份有限公司"],
        [7, "ONE-2026-550", "ONE APUS / 0022-019E", "TWKEL", "USOAK", "ONEU5678901", "40HC", "17200 kg", "2850", "2026-10-14", "2026-10-30", "PAID", "廣達電腦股份有限公司"],
        [8, "HMM-2026-773", "HMM ALGESIRAS / 0033-044W", "twkhh", "NLRTM", "HMMU6789012", "40HC", "20400", "USD 3,250", "15/10/2026", "2026/11/19", "collect", "台達電子工業股份有限公司"],
        [9, "  EMC-2026-009", "EVER GOLDEN / 0108-033W", "TWKHH", "USLAX", "EMCU7890123", "40GP", " 16,800 ", "$2,400", "2026/10/16", "2026-11-02", "PAID", "巨大機械工業股份有限公司"],
        [10, "ZIM-2026-118", "ZIM SAMMY OFER / 0012-005E", "TWKEL", "USNYC", "ZIMU8901234", "40HC", "23100", "3,800 USD", "2026-10-18", "2026-11-12", "Collect", "友達光電股份有限公司"],
        [11, "WHL-2026-150", "WAN HAI 502 / 0110-045S", "twkhh", "SGSIN", "WHLU1122334", "20GP", "11500 kg", "USD 950", "2026/10/19", "2026/10/25", "PAID", "台灣水泥股份有限公司"],
        [12, "MSK-2026-991", "MAERSK MC-KINNEY / 0411-002W", "TWKEL", "DEHAM", "MSKU3456789", "40HC", "22500", "3400 USD", "2026.10.12", "2026/11/15", "PAID", "日月光半導體製造股份有限公司"], # Duplicate
        [13, "YML-2026-204", "YM WISH / 0098-023E", "TWKEL", "USLGB", "YMLU2233445", "40HC", "19100", "$2,550", "2026-10-21", "2026-11-07", "PREPAID", "和碩聯合科技股份有限公司"],
        [14, "OOCL-2026-441", "OOCL GERMANY / 0048-019W", "twkhh", "BEANR", "OOLU3344556", "40HC", " 21,900 ", "3,300 USD", "22-10-2026", "2026/11/26", "Collect", "中國鋼鐵股份有限公司"],
        [15, "EMC-2026-035", "EVER BRAVE / 0067-088N", "TWKEL", "JPYOK", "EMCU4455667", "20GP", "8900 kg", "$800", "2026/10/23", "2026-10-27", "PAID", ""], # Missing shipper
        [16, "CMA-2026-580", "CMA CGM PALAIS / 0122-099W", "TWKHH", "FRFOS", "CMAU5566778", "40GP", "18200", "USD 2,950", "2026.10.24", "2026.11.28", "paid", "聯發科技股份有限公司"],
        [17, "ONE-2026-620", "ONE CYGNUS / 0015-008E", "twkhh", "USSEA", "ONEU6677889", "40HC", "17600", "2700 USD", "2026/10/25", "2026-11-10", "Collect", "緯創資通股份有限公司"],
        [18, "HMM-2026-890", "HMM OSLO / 0021-031W", "TWKEL", "GBFLX", "HMMU7788990", "40HC", "22800 kg", "$3,500", "2026-10-26", "2026/11/30", "PREPAID", "仁寶電腦工業股份有限公司"],
        [19, "WHL-2026-211", "WAN HAI 288 / 0077-015S", "twkhh", "MYPKG", "WHLU8899001", "20GP", " 10,200 ", "USD 1,050", "2026/10/27", "2026-11-03", "PAID", "統一企業股份有限公司"],
        [20, "YML-2026-315", "YM WONDER / 0112-064W", "TWKHH", "DEHAM", "YMLU9900112", "40HC", "20700", "3,350 USD", "28-10-2026", "2026/12/02", "", "正新橡膠工業股份有限公司"], # Missing payment
        [21, "MSK-2026-108", "MAERSK MADRID / 0205-011E", "twkel", "USLAX", "MSKU0011223", "40HC", "19500 kg", "$2,500", "2026.10.29", "2026-11-15", "PAID", "成霖企業股份有限公司"],
        [22, "EMC-2026-052", "EVER GENIUS / 0133-049W", "TWKHH", "NLRTM", "EMCU1133557", "40HC", "21400", "USD 3,200", "2026/10/30", "2026/12/04", "collect", "建大工業股份有限公司"],
        [23, "ZIM-2026-234", "ZIM MOUNT EVEREST / 0009-002E", "TWKEL", "USSAV", "ZIMU2244668", "40GP", " 17,900 ", "3,650", "2026-10-31", "2026-11-26", "PREPAID", "豐泰企業股份有限公司"],
        [24, "CMA-2026-671", "CMA CGM MARCO POLO / 0341-102W", "twkhh", "ITGOA", "CMAU3355779", "20GP", "9300", "$1,600", "2026/11/01", "2026-12-05", "PAID", "寶成工業股份有限公司"],
        [25, "OOCL-2026-590", "OOCL SCANDINAVIA / 0039-015W", "TWKEL", "GBSOU", "OOLU4466880", "40HC", "22100 kg", "USD 3,450", "02/11/2026", "2026/12/07", "Collect", "光寶科技股份有限公司"],
        [26, "WHL-2026-289", "WAN HAI 165 / 0045-009S", "twkhh", "THBKK", "WHLU5577991", "20GP", "10800", "$850", "2026-11-03", "2026/11/09", "paid", "英業達股份有限公司"],
        [27, "ONE-2026-733", "ONE COLUMBA / 0019-012E", "TWKHH", "USOAK", "ONEU6688002", "40HC", " 18,300 ", "2,800 USD", "2026/11/04", "2026-11-20", "PREPAID", "研華股份有限公司"],
        [28, "HMM-2026-945", "HMM DUBLIN / 0018-028W", "TWKEL", "BEANR", "HMMU7799113", "40HC", "21700 kg", "USD 3,300", "2026.11.05", "2026.12.10", "PAID", "台勝高新材料股份有限公司"],
        [29, "YML-2026-420", "YM WORLD / 0145-091W", "twkhh", "ESBCN", "YMLU8800224", "40GP", "16500", "$2,900", "06-11-2026", "2026/12/12", "collect", "致茂電子股份有限公司"],
        [30, "EMC-2026-077", "EVER GIFTED / 0121-040W", "TWKHH", "DEHAM", "EMCU9911335", "40HC", "20900", "USD 3,380", "2026/11/07", "2026-12-11", "PAID", "台達電子工業股份有限公司"]
    ]
    
    for r in raw_data_30:
        ws_raw.append(r)
    
    style_header_row(ws_raw, len(headers_raw), fill_color="800000") # Maroon header indicating raw/uncleaned
    style_data_rows(ws_raw, len(raw_data_30), len(headers_raw))
    autofit_columns(ws_raw)
    
    # ----------------- Sheet 2: Cleaned Master Table -----------------
    ws_clean = wb.create_sheet(title="航運標準資料_已清洗(Cleaned)")
    headers_clean = [
        "項次", "提單號碼", "船名", "航次編號", "裝貨港代碼", "卸貨港代碼", 
        "貨櫃號碼", "櫃型尺寸", "貨物毛重(KG)", "海運費(USD)", 
        "預計開航日(ETD)", "預計抵達日(ETA)", "航程天數(Transit Days)", "運費付款狀態", "託運人名稱"
    ]
    ws_clean.append(headers_clean)
    
    # Cleaned rows (30 -> 29 after removing duplicate row 12)
    cleaned_rows = [
        [1, "EMC-2026-001", "EVER GIVEN", "0142-052W", "TWKHH", "USLAX", "EMCU1234567", "40HC", 18500, 2450.00, "2026-10-05", "2026-10-22", 17, "PREPAID", "宏達國際貿易股份有限公司"],
        [2, "WHL-2026-088", "WAN HAI 315", "0089-012N", "TWKEL", "CNSHA", "WHLU9876543", "20GP", 12400, 1200.00, "2026-10-06", "2026-10-10", 4, "COLLECT", "台灣造紙工業股份有限公司"],
        [3, "YML-2026-105", "YM WELLBEING", "0214-088E", "TWKHH", "NLRTM", "YMLU4567890", "40GP", 21300, 3100.00, "2026-10-10", "2026-11-08", 29, "PREPAID", "長興材料工業股份有限公司"],
        [4, "OOCL-2026-302", "OOCL POLAND", "0055-031W", "TWKHH", "USLGB", "OOLU2345678", "40HC", 19800, 2600.00, "2026-10-11", "2026-10-28", 17, "PREPAID", "和大工業股份有限公司"],
        [5, "MSK-2026-991", "MAERSK MC-KINNEY", "0411-002W", "TWKEL", "DEHAM", "MSKU3456789", "40HC", 22500, 3400.00, "2026-10-12", "2026-11-15", 34, "PREPAID", "日月光半導體製造股份有限公司"],
        [6, "CMA-2026-412", "CMA CGM ANTOINE", "0918-077E", "TWKHH", "FRLEH", "CMAU8765432", "20GP", 9800, 1450.00, "2026-10-13", "2026-11-18", 36, "COLLECT", "台灣塑膠工業股份有限公司"],
        [7, "ONE-2026-550", "ONE APUS", "0022-019E", "TWKEL", "USOAK", "ONEU5678901", "40HC", 17200, 2850.00, "2026-10-14", "2026-10-30", 16, "PREPAID", "廣達電腦股份有限公司"],
        [8, "HMM-2026-773", "HMM ALGESIRAS", "0033-044W", "TWKHH", "NLRTM", "HMMU6789012", "40HC", 20400, 3250.00, "2026-10-15", "2026-11-19", 35, "COLLECT", "台達電子工業股份有限公司"],
        [9, "EMC-2026-009", "EVER GOLDEN", "0108-033W", "TWKHH", "USLAX", "EMCU7890123", "40GP", 16800, 2400.00, "2026-10-16", "2026-11-02", 17, "PREPAID", "巨大機械工業股份有限公司"],
        [10, "ZIM-2026-118", "ZIM SAMMY OFER", "0012-005E", "TWKEL", "USNYC", "ZIMU8901234", "40HC", 23100, 3800.00, "2026-10-18", "2026-11-12", 25, "COLLECT", "友達光電股份有限公司"],
        [11, "WHL-2026-150", "WAN HAI 502", "0110-045S", "TWKHH", "SGSIN", "WHLU1122334", "20GP", 11500, 950.00, "2026-10-19", "2026-10-25", 6, "PREPAID", "台灣水泥股份有限公司"],
        [12, "YML-2026-204", "YM WISH", "0098-023E", "TWKEL", "USLGB", "YMLU2233445", "40HC", 19100, 2550.00, "2026-10-21", "2026-11-07", 17, "PREPAID", "和碩聯合科技股份有限公司"],
        [13, "OOCL-2026-441", "OOCL GERMANY", "0048-019W", "TWKHH", "BEANR", "OOLU3344556", "40HC", 21900, 3300.00, "2026-10-22", "2026-11-26", 35, "COLLECT", "中國鋼鐵股份有限公司"],
        [14, "EMC-2026-035", "EVER BRAVE", "0067-088N", "TWKEL", "JPYOK", "EMCU4455667", "20GP", 8900, 800.00, "2026-10-23", "2026-10-27", 4, "PREPAID", "[待補託運人]"],
        [15, "CMA-2026-580", "CMA CGM PALAIS", "0122-099W", "TWKHH", "FRFOS", "CMAU5566778", "40GP", 18200, 2950.00, "2026-10-24", "2026-11-28", 35, "PREPAID", "聯發科技股份有限公司"],
        [16, "ONE-2026-620", "ONE CYGNUS", "0015-008E", "TWKHH", "USSEA", "ONEU6677889", "40HC", 17600, 2700.00, "2026-10-25", "2026-11-10", 16, "COLLECT", "緯創資通股份有限公司"],
        [17, "HMM-2026-890", "HMM OSLO", "0021-031W", "TWKEL", "GBFLX", "HMMU7788990", "40HC", 22800, 3500.00, "2026-10-26", "2026-11-30", 35, "PREPAID", "仁寶電腦工業股份有限公司"],
        [18, "WHL-2026-211", "WAN HAI 288", "0077-015S", "TWKHH", "MYPKG", "WHLU8899001", "20GP", 10200, 1050.00, "2026-10-27", "2026-11-03", 7, "PREPAID", "統一企業股份有限公司"],
        [19, "YML-2026-315", "YM WONDER", "0112-064W", "TWKHH", "DEHAM", "YMLU9900112", "40HC", 20700, 3350.00, "2026-10-28", "2026-12-02", 35, "[待確認款項]", "正新橡膠工業股份有限公司"],
        [20, "MSK-2026-108", "MAERSK MADRID", "0205-011E", "TWKEL", "USLAX", "MSKU0011223", "40HC", 19500, 2500.00, "2026-10-29", "2026-11-15", 17, "PREPAID", "成霖企業股份有限公司"],
        [21, "EMC-2026-052", "EVER GENIUS", "0133-049W", "TWKHH", "NLRTM", "EMCU1133557", "40HC", 21400, 3200.00, "2026-10-30", "2026-12-04", 35, "COLLECT", "建大工業股份有限公司"],
        [22, "ZIM-2026-234", "ZIM MOUNT EVEREST", "0009-002E", "TWKEL", "USSAV", "ZIMU2244668", "40GP", 17900, 3650.00, "2026-10-31", "2026-11-26", 26, "PREPAID", "豐泰企業股份有限公司"],
        [23, "CMA-2026-671", "CMA CGM MARCO POLO", "0341-102W", "TWKHH", "ITGOA", "CMAU3355779", "20GP", 9300, 1600.00, "2026-11-01", "2026-12-05", 34, "PREPAID", "寶成工業股份有限公司"],
        [24, "OOCL-2026-590", "OOCL SCANDINAVIA", "0039-015W", "TWKEL", "GBSOU", "OOLU4466880", "40HC", 22100, 3450.00, "2026-11-02", "2026-12-07", 35, "COLLECT", "光寶科技股份有限公司"],
        [25, "WHL-2026-289", "WAN HAI 165", "0045-009S", "TWKHH", "THBKK", "WHLU5577991", "20GP", 10800, 850.00, "2026-11-03", "2026-11-09", 6, "PREPAID", "英業達股份有限公司"],
        [26, "ONE-2026-733", "ONE COLUMBA", "0019-012E", "TWKHH", "USOAK", "ONEU6688002", "40HC", 18300, 2800.00, "2026-11-04", "2026-11-20", 16, "PREPAID", "研華股份有限公司"],
        [27, "HMM-2026-945", "HMM DUBLIN", "0018-028W", "TWKEL", "BEANR", "HMMU7799113", "40HC", 21700, 3300.00, "2026-11-05", "2026-12-10", 35, "PREPAID", "台勝高新材料股份有限公司"],
        [28, "YML-2026-420", "YM WORLD", "0145-091W", "TWKHH", "ESBCN", "YMLU8800224", "40GP", 16500, 2900.00, "2026-11-06", "2026-12-12", 36, "COLLECT", "致茂電子股份有限公司"],
        [29, "EMC-2026-077", "EVER GIFTED", "0121-040W", "TWKHH", "DEHAM", "EMCU9911335", "40HC", 20900, 3380.00, "2026-11-07", "2026-12-11", 34, "PREPAID", "台達電子工業股份有限公司"]
    ]
    for r in cleaned_rows:
        ws_clean.append(r)
        
    style_header_row(ws_clean, len(headers_clean), fill_color="1F497D") # Navy blue header for cleaned
    style_data_rows(ws_clean, len(cleaned_rows), len(headers_clean))
    autofit_columns(ws_clean)
    
    # ----------------- Sheet 3: Cleaning Formula & Rule Mapping -----------------
    ws_rules = wb.create_sheet(title="清洗函數與操作對照表(Rules)")
    headers_rules = ["清洗目標欄位", "原始資料常見缺陷", "推薦Excel功能選項(UI)", "推薦Excel函數公式(Formula)", "說明與實務注意事項"]
    ws_rules.append(headers_rules)
    
    rule_data = [
        ["提單號碼", "前後夾帶不可見空白字元", "常用 -> 尋找與選取 -> 取代 (或資料剖析)", "=TRIM(B2)", "TRIM清除前後空格，但保留單一間隔空格。"],
        ["船名航次", "船名與航次合併在同欄", "資料 -> 資料剖析 (以'/'分隔) 或 快速填入(Ctrl+E)", "=TRIM(LEFT(C2, FIND(\"/\", C2)-1))\n=TRIM(MID(C2, FIND(\"/\", C2)+1, 50))", "航運系統產出常將Vessel與Voyage擠在同一欄位，拆解後方能統計個別船舶趟次。"],
        ["裝卸港代碼", "大小寫紊亂(twkhh vs TWKHH)", "無直接UI按鈕(需靠公式或Power Query)", "=UPPER(TRIM(D2))", "港口代碼必須為標準5碼大寫UN/LOCODE，便於後續VLOOKUP/XLOOKUP關聯港口名稱。"],
        ["貨櫃號碼", "尾部有空格、混雜小寫", "常用 -> 快速填入 (Ctrl+E)", "=UPPER(TRIM(F2))", "國際貨櫃編號標準(ISO 6346)為4碼英文+7碼數字，多餘空格會導致EDI報單比對失敗。"],
        ["貨物毛重", "帶單位'kg'、千分位逗號、字串格式", "資料 -> 資料剖析 (自動轉數值) 或 取代", "=VALUE(SUBSTITUTE(SUBSTITUTE(LOWER(H2),\"kg\",\"\"),\" \",\"\"))", "數值若為文字格式無法進行SUM或樞紐分析，必須轉成實體數值型態(Number)。"],
        ["海運費", "混雜幣別代碼($、USD、EUR、空格)", "常用 -> 取代 (將USD/EUR/$替換為空)", "=VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(I2,\"$\",\"\"),\"USD\",\"\"),\"EUR\",\"\"))", "運費計算前必須先將貨幣符號分離，幣別獨立一欄，金額轉為純數字。"],
        ["預計開航/抵達日", "多種日期格式混雜(/、-、.)、字串型態", "資料 -> 資料剖析 (第3步指定日期格式YMD)", "=DATEVALUE(SUBSTITUTE(SUBSTITUTE(J2,\".\",\"/\"),\"-\",\"/\"))", "統一為標準Excel日期序列值(如YYYY-MM-DD)，才能直接相減計算航程天數(ETA-ETD)。"],
        ["重複紀錄", "同一提單重複建檔(Row 5 & 12)", "資料 -> 移除重複項 (勾選提單號碼)", "=IF(COUNTIF($B$2:B2, B2)>1, \"重複\", \"正常\")", "移除重複是資料分析前最關鍵的常態檢驗，避免運費與貨量雙重計算。"],
        ["遺漏值", "託運人或付款狀態為空白", "常用 -> 尋找與選取 -> 特殊目標 -> 空格", "=IF(M2=\"\", \"[待補託運人]\", M2)", "空白儲存格在資料分析中容易造成分類失真，應填入業務標準預設字樣。"]
    ]
    for r in rule_data:
        ws_rules.append(r)
        
    style_header_row(ws_rules, len(headers_rules), fill_color="2E75B6")
    style_data_rows(ws_rules, len(rule_data), len(headers_rules))
    autofit_columns(ws_rules, max_len_cap=50)
    
    wb.save("Shipping_Demo_30.xlsx")
    print("Saved Shipping_Demo_30.xlsx successfully.")

# ==========================================
# 2. GENERATE 300-ROW PRACTICE DATASET
# ==========================================
def generate_300_practice():
    wb = openpyxl.Workbook()
    
    # ----------------- Sheet 1: Practice Raw 300 -----------------
    ws_raw = wb.active
    ws_raw.title = "航運實務情境練習檔(Raw_300)"
    
    headers_300 = [
        "序號", "訂艙編號(Booking No)", "提單號碼(B/L No)", "船名(Vessel)", "航次(Voyage)", 
        "起運港(POL)", "目的港(POD)", "貨櫃編號(Container No)", "櫃型代碼(ISO Type)", 
        "貨品類別(Commodity)", "件數(Packages)", "包裝單位", "毛重(Gross Weight)", 
        "體積(CBM)", "海運費幣別", "單價運費(Rate)", "其他附加費(Surcharges)", 
        "預計離港日(ETD)", "預計到港日(ETA)", "實際到港日(ATA)", "提櫃日(Gate-Out)", 
        "免租期天數(Free Days)", "付款條件(Terms)", "海關放行狀態(Customs Status)", "客戶名稱(Customer)"
    ]
    ws_raw.append(headers_300)
    
    # Pre-defined pools for realistic generation
    vessels = [
        ("EVER GIVEN", "0142-052W"), ("EVER GLORY", "0088-034E"), ("EVER ACE", "0012-005W"),
        ("WAN HAI 315", "0089-012N"), ("WAN HAI 502", "0110-045S"), ("WAN HAI 288", "0077-015S"),
        ("YM WELLBEING", "0214-088E"), ("YM WISH", "0098-023E"), ("YM WONDER", "0112-064W"),
        ("OOCL POLAND", "0055-031W"), ("OOCL GERMANY", "0048-019W"), ("OOCL SCANDINAVIA", "0039-015W"),
        ("MAERSK MC-KINNEY", "0411-002W"), ("MAERSK MADRID", "0205-011E"), ("MAERSK MC-KINNEY", "0411-002W"),
        ("CMA CGM ANTOINE", "0918-077E"), ("CMA CGM PALAIS", "0122-099W"), ("CMA CGM MARCO POLO", "0341-102W"),
        ("ONE APUS", "0022-019E"), ("ONE CYGNUS", "0015-008E"), ("ONE COLUMBA", "0019-012E"),
        ("HMM ALGESIRAS", "0033-044W"), ("HMM OSLO", "0021-031W"), ("HMM DUBLIN", "0018-028W")
    ]
    
    pols = ["TWKHH", "TWKEL", "TWTXG", "twkhh", "twkel", " TWKHH ", "Twtxg"]
    pods = ["USLAX", "USLGB", "USOAK", "USNYC", "DEHAM", "NLRTM", "BEANR", "GBFLX", "JPYOK", "CNSHA", "SGSIN", "THBKK", "MYPKG", "ITGOA"]
    container_types = ["20GP", "40GP", "40HC", "45HC", "20RF", "40RH", "20OT", "40FR"]
    commodities = ["電子零組件", "精密機械", "自行車成車及零件", "紡織成衣", "塑膠原料", "車用晶片", "化學原料", "冷凍水產", "金屬五金", "辦公家具"]
    pkg_units = ["CTN", "PLT", "BAG", "PKG", "DRUM", "CRT", "ctn", "plt"]
    currencies = ["USD", "EUR", "TWD"]
    terms = ["Prepaid", "Collect", "PREPAID", "COLLECT", " prepaid ", "collect"]
    customs = ["CLEARED", "HOLD", "INSPECTION", "PENDING", "cleared", "hold", ""]
    customers = [
        "台積電製造股份有限公司", "鴻海精密工業股份有限公司", "聯發科技股份有限公司", "廣達電腦股份有限公司", 
        "台達電子工業股份有限公司", "巨大機械工業股份有限公司", "美麗達工業股份有限公司", "正新橡膠工業股份有限公司", 
        "和大工業股份有限公司", "和碩聯合科技股份有限公司", "友達光電股份有限公司", "群創光電股份有限公司", 
        "長興材料工業股份有限公司", "台灣塑膠工業股份有限公司", "中國鋼鐵股份有限公司", "宏達國際貿易股份有限公司",
        "研華股份有限公司", "英業達股份有限公司", "緯創資通股份有限公司", "台勝高新材料股份有限公司"
    ]
    
    random.seed(42) # Deterministic generation for consistency
    base_date = datetime(2026, 10, 1)
    
    raw_rows = []
    # Inject deliberate duplicates, spacing errors, date format glitches, and text numbers
    for i in range(1, 301):
        # Scenario variations
        carrier_prefix = random.choice(["EMC", "WHL", "YML", "OOL", "MSK", "CMA", "ONE", "HMM"])
        bkg_no = f"BKG-{carrier_prefix}-{20260000 + i}"
        bl_no = f"{carrier_prefix}TW{2600000 + i}"
        
        # Inject space error into some BLs
        if i % 7 == 0:
            bl_no = f"  {bl_no}  "
        elif i % 13 == 0:
            bl_no = bl_no.lower()
            
        ves, voy = random.choice(vessels)
        # Combine or add weird spacing
        if i % 5 == 0:
            ves_field = f"  {ves}  "
        else:
            ves_field = ves
            
        pol = random.choice(pols)
        pod = random.choice(pods)
        if i % 9 == 0:
            pod = pod.lower()
            
        # Container number: 4 letters + 7 digits
        owner_code = carrier_prefix[:3] + "U"
        cntr_no = f"{owner_code}{random.randint(1000000, 9999999)}"
        if i % 11 == 0:
            cntr_no = f"{cntr_no} " # trailing space
        elif i % 17 == 0:
            cntr_no = cntr_no.lower()
            
        iso_type = random.choice(container_types)
        comm = random.choice(commodities)
        pkgs = random.randint(50, 2500)
        unit = random.choice(pkg_units)
        
        # Weight & Volume
        raw_wt = random.randint(3500, 26000)
        if i % 4 == 0:
            gross_wt = f"{raw_wt:,} kg"
        elif i % 6 == 0:
            gross_wt = f"{raw_wt} KGS"
        elif i % 8 == 0:
            gross_wt = f" {raw_wt} "
        else:
            gross_wt = str(raw_wt)
            
        cbm = round(random.uniform(12.5, 68.0), 2)
        curr = random.choice(currencies)
        raw_rate = random.choice([850, 1200, 1850, 2400, 2800, 3200, 3600, 4100])
        
        # Currency rate string formatting flaws
        if i % 3 == 0:
            rate_str = f"${raw_rate:,}"
        elif i % 5 == 0:
            rate_str = f"{curr} {raw_rate}"
        elif i % 7 == 0:
            rate_str = f"{raw_rate} USD"
        else:
            rate_str = str(raw_rate)
            
        surcharges = round(random.uniform(150, 650), 2)
        
        # Dates
        etd_dt = base_date + timedelta(days=int(i * 0.4) + random.randint(0, 3))
        transit = random.randint(4, 38)
        eta_dt = etd_dt + timedelta(days=transit)
        ata_dt = eta_dt + timedelta(days=random.choice([-1, 0, 0, 1, 2, 4])) # ATA might be delayed
        free_days = random.choice([7, 10, 14, 21])
        # Gate out date: mostly within free days, but some exceed free days (demurrage occurred!)
        gate_out_delay = random.choice([2, 5, 6, 7, 8, 12, 16, 25])
        gate_out_dt = ata_dt + timedelta(days=gate_out_delay)
        
        # Format dates with inconsistency
        if i % 4 == 0:
            etd_str = etd_dt.strftime("%Y/%m/%d")
            eta_str = eta_dt.strftime("%Y/%m/%d")
            ata_str = ata_dt.strftime("%Y/%m/%d")
            gate_str = gate_out_dt.strftime("%Y/%m/%d")
        elif i % 5 == 0:
            etd_str = etd_dt.strftime("%Y-%m-%d")
            eta_str = eta_dt.strftime("%Y-%m-%d")
            ata_str = ata_dt.strftime("%Y-%m-%d")
            gate_str = gate_out_dt.strftime("%Y-%m-%d")
        elif i % 6 == 0:
            etd_str = etd_dt.strftime("%d/%m/%Y")
            eta_str = eta_dt.strftime("%d/%m/%Y")
            ata_str = ata_dt.strftime("%d/%m/%Y")
            gate_str = gate_out_dt.strftime("%d/%m/%Y")
        else:
            etd_str = etd_dt.strftime("%Y.%m.%d")
            eta_str = eta_dt.strftime("%Y.%m.%d")
            ata_str = ata_dt.strftime("%Y.%m.%d")
            gate_str = gate_out_dt.strftime("%Y.%m.%d")
            
        term = random.choice(terms)
        cust_status = random.choice(customs)
        cust_name = random.choice(customers)
        if i % 29 == 0:
            cust_name = "" # missing customer name
            
        row = [
            i, bkg_no, bl_no, ves_field, voy, pol, pod, cntr_no, iso_type, comm,
            pkgs, unit, gross_wt, cbm, curr, rate_str, surcharges,
            etd_str, eta_str, ata_str, gate_str, free_days, term, cust_status, cust_name
        ]
        raw_rows.append(row)
        
    # Inject exact duplicates for practice
    raw_rows[45] = list(raw_rows[10]) # duplicate row 11 into 46
    raw_rows[45][0] = 46 # update seq num
    raw_rows[180] = list(raw_rows[75])
    raw_rows[180][0] = 181
    raw_rows[260] = list(raw_rows[120])
    raw_rows[260][0] = 261
    
    for r in raw_rows:
        ws_raw.append(r)
        
    style_header_row(ws_raw, len(headers_300), fill_color="371435") # Plum header
    style_data_rows(ws_raw, len(raw_rows), len(headers_300))
    autofit_columns(ws_raw)
    
    # ----------------- Sheet 2: Reference Port Lookup -----------------
    ws_port = wb.create_sheet(title="港口代碼對照表(Port_Master)")
    port_headers = ["港口代碼(UN/LOCODE)", "港口中文名稱", "港口英文名稱", "所屬國家/地區", "主要航線區域"]
    ws_port.append(port_headers)
    port_master = [
        ["TWKHH", "高雄港", "Kaohsiung", "台灣 (Taiwan)", "亞洲/台灣"],
        ["TWKEL", "基隆港", "Keelung", "台灣 (Taiwan)", "亞洲/台灣"],
        ["TWTXG", "台中港", "Taichung", "台灣 (Taiwan)", "亞洲/台灣"],
        ["USLAX", "洛杉磯港", "Los Angeles", "美國 (USA)", "美西航線(USWC)"],
        ["USLGB", "長堤港", "Long Beach", "美國 (USA)", "美西航線(USWC)"],
        ["USOAK", "奧克蘭港", "Oakland", "美國 (USA)", "美西航線(USWC)"],
        ["USNYC", "紐約港", "New York", "美國 (USA)", "美東航線(USEC)"],
        ["DEHAM", "漢堡港", "Hamburg", "德國 (Germany)", "歐洲線(Europe)"],
        ["NLRTM", "鹿特丹港", "Rotterdam", "荷蘭 (Netherlands)", "歐洲線(Europe)"],
        ["BEANR", "安特衛普港", "Antwerp", "比利時 (Belgium)", "歐洲線(Europe)"],
        ["GBFLX", "費利克斯托港", "Felixstowe", "英國 (UK)", "歐洲線(Europe)"],
        ["JPYOK", "橫濱港", "Yokohama", "日本 (Japan)", "東北亞線(Northeast Asia)"],
        ["CNSHA", "上海港", "Shanghai", "中國 (China)", "兩岸/亞洲線(Intra-Asia)"],
        ["SGSIN", "新加坡港", "Singapore", "新加坡 (Singapore)", "東南亞轉運樞紐"],
        ["THBKK", "曼谷港", "Bangkok", "泰國 (Thailand)", "東南亞線(Southeast Asia)"],
        ["MYPKG", "巴生港", "Port Klang", "馬來西亞 (Malaysia)", "東南亞線(Southeast Asia)"],
        ["ITGOA", "熱那亞港", "Genoa", "義大利 (Italy)", "地中海線(Mediterranean)"]
    ]
    for p in port_master:
        ws_port.append(p)
    style_header_row(ws_port, len(port_headers), fill_color="1F497D")
    style_data_rows(ws_port, len(port_master), len(port_headers))
    autofit_columns(ws_port)
    
    # ----------------- Sheet 3: ISO Container Type Lookup -----------------
    ws_iso = wb.create_sheet(title="櫃型代碼對照表(ISO_Type)")
    iso_headers = ["櫃型代碼", "長度(呎)", "高度規格", "櫃型分類中文說明", "TEU換算係數", "最大容積(CBM)", "適用貨物特性"]
    ws_iso.append(iso_headers)
    iso_master = [
        ["20GP", "20 FT", "8'6\" 標準", "20呎普通乾貨櫃 (General Purpose)", 1.0, 33.2, "重貨、金屬五金、化學原料、礦石"],
        ["40GP", "40 FT", "8'6\" 標準", "40呎普通乾貨櫃 (General Purpose)", 2.0, 67.7, "一般體積型貨物、家具、紡織品"],
        ["40HC", "40 FT", "9'6\" 高櫃", "40呎超高乾貨櫃 (High Cube)", 2.0, 76.4, "輕泡貨、電子產品、腳踏車、體積型電器"],
        ["45HC", "45 FT", "9'6\" 超高", "45呎超長超高櫃 (High Cube)", 2.25, 86.0, "巨量輕泡貨、跨國大型量販貨物"],
        ["20RF", "20 FT", "8'6\" 標準", "20呎冷凍冷藏櫃 (Reefer)", 1.0, 28.3, "小型冷凍肉品、高價生鮮蔬果、特殊藥品"],
        ["40RH", "40 FT", "9'6\" 高櫃", "40呎超高冷凍櫃 (Reefer High Cube)", 2.0, 67.8, "大宗冷凍水產、肉品、乳製品、生鮮溫控"],
        ["20OT", "20 FT", "開頂", "20呎開頂櫃 (Open Top)", 1.0, 32.0, "超高機具、頂部吊裝設備、玻璃片"],
        ["40FR", "40 FT", "平板/折疊", "40呎平板櫃 (Flat Rack)", 2.0, 0.0, "超寬超高超重機械、遊艇、大型車輛"]
    ]
    for item in iso_master:
        ws_iso.append(item)
    style_header_row(ws_iso, len(iso_headers), fill_color="1F497D")
    style_data_rows(ws_iso, len(iso_master), len(iso_headers))
    autofit_columns(ws_iso)
    
    # ----------------- Sheet 4: Reference Cleaned Master & Calculated Fields -----------------
    ws_sol = wb.create_sheet(title="參考答案_標準清洗表(Cleaned_Master)")
    sol_headers = [
        "序號", "訂艙編號", "標準提單號", "船名", "航次", "起運港代碼", "起運港中文名", 
        "目的港代碼", "目的港中文名", "標準貨櫃號", "櫃型代碼", "換算TEU", "貨品類別", 
        "件數", "標準單位", "純毛重(KG)", "體積(CBM)", "幣別", "單價運費(USD)", "附加費", 
        "標準ETD", "標準ETA", "航程天數", "標準ATA", "標準提櫃日", "港口停留天數(Port Stay)", 
        "免租天數", "超期留港天數(Overdue)", "滯櫃警示(Demurrage Alert)", "付款條件", "海關狀態", "客戶名稱", "異常檢核標記"
    ]
    ws_sol.append(sol_headers)
    
    # Port map
    port_dict = {p[0]: p[1] for p in port_master}
    teu_dict = {i[0]: i[4] for i in iso_master}
    
    cleaned_count = 0
    seen_bls = set()
    
    for r in raw_rows:
        seq = r[0]
        bkg = r[1].strip()
        bl = r[2].strip().upper()
        
        # Check duplicate
        is_dup = bl in seen_bls
        seen_bls.add(bl)
        
        ves = r[3].strip()
        voy = r[4].strip()
        pol = r[5].strip().upper()
        pol_name = port_dict.get(pol, "未知港口")
        pod = r[6].strip().upper()
        pod_name = port_dict.get(pod, "未知港口")
        cntr = r[7].strip().upper()
        iso = r[8].strip().upper()
        teu = teu_dict.get(iso, 1.0)
        comm = r[9].strip()
        pkgs = int(r[10])
        unit = r[11].strip().upper()
        
        # Weight clean
        wt_str = str(r[12]).lower().replace("kg", "").replace("kgs", "").replace(",", "").strip()
        try:
            wt_val = float(wt_str)
        except:
            wt_val = 0.0
            
        cbm_val = float(r[13])
        curr = r[14].strip().upper()
        
        # Rate clean
        rate_str = str(r[15]).replace("$", "").replace("USD", "").replace("EUR", "").replace("TWD", "").replace(",", "").strip()
        try:
            rate_val = float(rate_str)
        except:
            rate_val = 0.0
            
        surch = float(r[16])
        
        # Standardize dates
        def parse_date_str(d_str):
            d_str = str(d_str).strip().replace(".", "/").replace("-", "/")
            parts = d_str.split("/")
            if len(parts) == 3:
                if len(parts[0]) == 4: # YYYY/MM/DD
                    y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
                else: # DD/MM/YYYY
                    d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
                return datetime(y, m, d)
            return datetime(2026, 10, 1)
            
        etd_d = parse_date_str(r[17])
        eta_d = parse_date_str(r[18])
        ata_d = parse_date_str(r[19])
        gate_d = parse_date_str(r[20])
        
        transit_days = (eta_d - etd_d).days
        port_stay = (gate_d - ata_d).days
        free_days = int(r[21])
        overdue_days = max(0, port_stay - free_days)
        demurrage_alert = "滯港罰款(Demurrage)" if overdue_days > 0 else "正常(Normal)"
        
        term = r[22].strip().upper()
        customs_st = r[23].strip().upper() if r[23].strip() else "PENDING"
        cust = r[24].strip() if r[24].strip() else "[待確認客戶]"
        
        audit_note = []
        if is_dup:
            audit_note.append("提單重複")
        if cust == "[待確認客戶]":
            audit_note.append("缺客戶名稱")
        if overdue_days > 0:
            audit_note.append(f"超期滯港{overdue_days}天")
        audit_str = " | ".join(audit_note) if audit_note else "檢核通過"
        
        sol_row = [
            seq, bkg, bl, ves, voy, pol, pol_name, pod, pod_name, cntr, iso, teu, comm,
            pkgs, unit, wt_val, cbm_val, curr, rate_val, surch,
            etd_d.strftime("%Y-%m-%d"), eta_d.strftime("%Y-%m-%d"), transit_days,
            ata_d.strftime("%Y-%m-%d"), gate_d.strftime("%Y-%m-%d"), port_stay,
            free_days, overdue_days, demurrage_alert, term, customs_st, cust, audit_str
        ]
        ws_sol.append(sol_row)
        cleaned_count += 1
        
    style_header_row(ws_sol, len(sol_headers), fill_color="1F497D")
    style_data_rows(ws_sol, cleaned_count, len(sol_headers))
    autofit_columns(ws_sol)
    
    # ----------------- Sheet 5: Exercise Instructions & Tasks -----------------
    ws_tasks = wb.create_sheet(title="綜合練習題目與操作要求(Tasks)")
    task_headers = ["題號", "作業名稱", "練習情境描述", "目標欄位與清洗條件", "使用建議之Excel工具/函數", "預期產出標準"]
    ws_tasks.append(task_headers)
    tasks = [
        ["任務一", "提單與訂艙編號標準化", "訂艙與提單編號來自不同貨代EDI，存在前後空格與大小寫紊亂問題。", "Booking No, B/L No", "TRIM, UPPER, 或 快速填入(Ctrl+E)", "所有編號消除任何前後空白，英文字母全數轉為大寫標準格式。"],
        ["任務二", "裝卸港UN/LOCODE與中文名關聯", "POL與POD欄位大小寫混雜，且需為行政人員補上中文港口全名以利報表閱讀。", "POL, POD -> 港口中文名", "UPPER + XLOOKUP / VLOOKUP 參照 Port_Master", "港口代碼全大寫，並正確帶出『高雄港』、『鹿特丹港』等對應中文港名。"],
        ["任務三", "貨櫃編號與ISO櫃型規格關聯", "貨櫃號碼需符合ISO 6346大寫無空格，並根據櫃型代碼換算TEU係數以利統計運載量。", "Container No, ISO Type -> TEU", "TRIM, UPPER, XLOOKUP 參照 ISO_Type", "櫃號大寫無空格，ISO代碼轉為標準TEU(20呎=1, 40呎=2, 45呎=2.25)。"],
        ["任務四", "貨物毛重與包裝單位數值化", "毛重欄位混雜'kg'、'KGS'與千分位符號，單位欄位大小寫混亂，導致無法加總。", "Gross Weight -> 數值, 包裝單位", "SUBSTITUTE, VALUE, UPPER, 資料剖析", "毛重轉為純數值型態(可加總)，包裝單位統一為大寫(CTN, PLT)。"],
        ["任務五", "海運費幣別分離與金額清洗", "單價運費混有貨幣符號($、USD)，無法進行財務乘算與樞紐分析。", "單價運費 -> 純數字, 幣別代碼", "尋找與取代, SUBSTITUTE, VALUE", "金額轉為純數字型態，幣別獨立欄位，可進行總運費試算(Rate + Surcharges)。"],
        ["任務六", "航程日期清洗與航期天數計算", "ETD與ETA混雜四種不同日期分隔符(/、-、.)及文字格式，需計算航程天數。", "ETD, ETA -> 航程天數", "資料剖析(指定YMD) 或 DATEVALUE/SUBSTITUTE, ETA-ETD", "日期轉換為標準日期序列值(YYYY-MM-DD)，計算出正確航程天數。"],
        ["任務七", "港口免租期與滯港費(Demurrage)預警", "根據實際到港日(ATA)、提櫃日(Gate-Out)與免租天數，計算是否產生超期留港費用。", "ATA, Gate-Out, Free Days -> 滯港天數與警示", "相減天數, IF, MAX 函數", "計算港口停留天數(Gate Out - ATA)，超期天數=MAX(0, 停留天數-免租期)，超期者標示『滯港罰款』。"],
        ["任務八", "全域重複提單排查與缺漏資料稽核", "排查300筆資料中是否存在重複建單，並填補遺漏的客戶名稱與海關狀態。", "整份工作表, B/L No, Customer, Customs", "資料 -> 移除重複項 或 COUNTIF條件格式設定, IF/ISBLANK", "標示出重複提單序號(如第46, 181, 261筆)，空客戶填寫『[待確認客戶]』，空海關填『PENDING』。"]
    ]
    for t in tasks:
        ws_tasks.append(t)
    style_header_row(ws_tasks, len(task_headers), fill_color="2E75B6")
    style_data_rows(ws_tasks, len(tasks), len(task_headers))
    autofit_columns(ws_tasks, max_len_cap=45)
    
    wb.save("Shipping_Practice_300.xlsx")
    print("Saved Shipping_Practice_300.xlsx successfully.")

if __name__ == "__main__":
    generate_30_demo()
    generate_300_practice()
