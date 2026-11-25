import base64
import os
import json

# 1. Load and base64-encode all images
b64_images = {}
for f in sorted(os.listdir("images")):
    if f.endswith(".png"):
        p = os.path.join("images", f)
        with open(p, "rb") as img_f:
            b64_images[f] = f"data:image/png;base64,{base64.b64encode(img_f.read()).decode('utf-8')}"

print(f"Loaded {len(b64_images)} base64 images.")

# Define Slide Content and Metadata
slides_data = [
    # Slide 1: Cover
    {
        "id": "slide-1",
        "chapter": "封面",
        "title": "航運行政人員 MS Excel 基礎資料前處理操作實戰指引",
        "subtitle": "從零認識介面、標準化資料衛生、30筆真實業務示範到300筆大數據案例全解",
        "badge": "MARITIME LOGISTICS & DATA ANALYTICS TRAINING",
        "content_type": "cover",
        "tips": "【演講者備忘】：歡迎各位航運物流同仁。本課程專注於解決大家在日常提單、訂艙與運費報表中遇到的『髒資料』問題。掌握前處理，將節省您 80% 的無效核對時間！"
    },
    # Slide 2: Agenda
    {
        "id": "slide-2",
        "chapter": "課程大綱",
        "title": "培訓課程全景地圖 (Agenda)",
        "subtitle": "循序漸進的五大實戰模組，打造航運數據標準化作業思維",
        "content_type": "agenda",
        "tips": "【演講者備忘】：課程由淺入深，從最基礎的介面巡禮，到資料衛生觀念，再透過 30 筆示範操作與 300 筆大數據挑戰，最後提供永久速查工具手冊。"
    },
    # Slide 3: Ch1 Why Preprocessing
    {
        "id": "slide-3",
        "chapter": "第一章：介面與新手導航",
        "title": "1.1 航運行政人員為何必須精通資料前處理？",
        "subtitle": "資料科學的 80/20 法則與『垃圾進，垃圾出 (GIGO)』效應",
        "content_type": "why_prep",
        "tips": "【實務警語】：航運提單多一個不可見空格，VLOOKUP 就抓不到報關行資訊；重量多打一個 'kg'，SUM 加總直接歸零！前處理是商業決策的唯一起點。"
    },
    # Slide 4: Ch1 Interface Overview
    {
        "id": "slide-4",
        "chapter": "第一章：介面與新手導航",
        "title": "1.2 Excel 核心操作環境八大功能區全景導覽",
        "subtitle": "標題列、功能區、名稱方塊、資料編輯列與狀態列的核心職能",
        "content_type": "image_focus",
        "image_key": "fig1_excel_interface.png",
        "caption": "圖 1-1：Excel 核心操作介面八大功能區導覽示意圖 (點擊圖片可放大檢視)",
        "tips": "【操作關鍵】：狀態列是免費的資料型態檢測器！框選一整欄數字，若底端只顯示『計數』卻沒有『加總』，代表這整欄數字被存成了文字型態！"
    },
    # Slide 5: Ch1 Shortcuts
    {
        "id": "slide-5",
        "chapter": "第一章：介面與新手導航",
        "title": "1.3 航運行政人員必備之十組超高效快捷鍵",
        "subtitle": "擺脫滑鼠拉拉拉的手痠夢魘，操作效率瞬間飆升 5 倍",
        "content_type": "shortcuts",
        "tips": "【操作訣竅】：最神快捷鍵首推【Ctrl + Shift + L】(一秒開關自動篩選) 與【Ctrl + E】(快速填入智慧萃取)！"
    },
    # Slide 6: Ch2 Data Hygiene & 1NF
    {
        "id": "slide-6",
        "chapter": "第二章：資料衛生與基本常識",
        "title": "2.1 資料衛生 (Data Hygiene) 與第一正規化 (1NF)",
        "subtitle": "拒絕報表排版型錯誤結構，回歸扁平化分析型標準二維表",
        "content_type": "image_focus",
        "image_key": "fig2_data_hygiene_good_vs_bad.png",
        "caption": "圖 2-1：報表排版型錯誤結構 vs 扁平化分析型標準二維表對比圖 (點擊圖片可放大)",
        "tips": "【鐵律宣導】：資料明細區域『絕對嚴禁使用跨欄置中/合併儲存格』！合併儲存格會讓電腦底層產生空值，徹底癱瘓篩選與樞紐分析！"
    },
    # Slide 7: Ch2 7 Dirty Data Diseases
    {
        "id": "slide-7",
        "chapter": "第二章：資料衛生與基本常識",
        "title": "2.2 航運行政常見『七大骯髒資料 (Dirty Data)』互動診斷盤",
        "subtitle": "點擊各項病症卡片，查看其破壞性後果與標準處理策略",
        "content_type": "interactive_dirty_data",
        "tips": "【演講者提示】：請學員點選卡片互動，了解各項骯髒資料在海運實務中如何導致 EDI 傳輸失敗與運費短漏。"
    },
    # Slide 8: Ch2 4 Data Types & Traps
    {
        "id": "slide-8",
        "chapter": "第二章：資料衛生與基本常識",
        "title": "2.3 Excel 四大本質資料型態與隱形陷阱",
        "subtitle": "看清本質而非外觀：文字、數值、日期序列值與綠色錯誤三角形",
        "content_type": "data_types",
        "tips": "【新手避坑】：合格的日期在 Excel 內部是自 1900/1/1 起算的天數數值！只有真正的日期序列值，才能直接相減 (ETA - ETD) 計算航程天數。"
    },
    # Slide 9: Ch3 30-Row Demo Overview
    {
        "id": "slide-9",
        "chapter": "第三章：30筆實戰示範",
        "title": "3.1 30 筆真實航運業務資料檔：清洗前後對比",
        "subtitle": "涵蓋遠洋與近洋提單、包含前後空格、混亂代碼、多樣日期與重複提單",
        "content_type": "interactive_table_toggle",
        "tips": "【實例說明】：切換上方頁籤，觀察原本帶有單位、混亂日期的原始表格，如何被梳理成標準化的一維資料庫。"
    },
    # Slide 10: Ch3 Text to Columns
    {
        "id": "slide-10",
        "chapter": "第三章：30筆實戰示範",
        "title": "3.2 功能選項 1：【資料剖析】3 步驟精靈拆解船名航次",
        "subtitle": "以斜線 '/' 為界，將擠在同格的屬性無損原子化分割",
        "content_type": "image_focus",
        "image_key": "fig4_text_to_columns_wizard.png",
        "caption": "圖 3-2：【資料剖析】精靈 3 步驟實戰分解示意圖 (點擊圖片可放大)",
        "tips": "【重要細節】：步驟 3 的『目的端』預設為覆蓋原始欄位，請務必手動修改為新空白欄位 (如 $D$2)，保留原始軌跡！"
    },
    # Slide 11: Ch3 Flash Fill Ctrl+E
    {
        "id": "slide-11",
        "chapter": "第三章：30筆實戰示範",
        "title": "3.3 功能選項 2：【快速填入 (Ctrl+E)】智慧萃取黑科技",
        "subtitle": "手動打一筆示範，按下 Ctrl+E 瞬間完成全欄 30 筆拆解！",
        "content_type": "image_focus",
        "image_key": "fig3_flash_fill_ctrl_e.png",
        "caption": "圖 3-1：Excel 快速填入 (Flash Fill) 快捷鍵【Ctrl + E】操作拆解 (點擊圖片可放大)",
        "tips": "【口訣記憶】：一建立範本、二游標下移、三按下【Ctrl + E】！快速填入是微軟最強大的字串拆解捷徑。"
    },
    # Slide 12: Ch3 Remove Duplicates
    {
        "id": "slide-12",
        "chapter": "第三章：30筆實戰示範",
        "title": "3.4 功能選項 3：【移除重複項】排除重複提單",
        "subtitle": "以提單號碼 (B/L No) 為核心主鍵，避免運費與貨量雙重虛胖計算",
        "content_type": "image_focus",
        "image_key": "fig5_remove_duplicates_dialog.png",
        "caption": "圖 3-3：【移除重複項】對話方塊設定與提單主鍵排重示範 (點擊圖片可放大)",
        "tips": "【防錯技巧】：點擊『取消全選』，單獨勾選『提單號碼』！只要提單重複就代表同一票貨重複建檔，成功抓出 Row 12 重複單據。"
    },
    # Slide 13: Ch3 Find/Replace & Validation
    {
        "id": "slide-13",
        "chapter": "第三章：30筆實戰示範",
        "title": "3.5 功能選項 4 & 5：【尋找與取代】與【資料驗證】源頭防錯",
        "subtitle": "批次清除 'kg' 單位與 '$' 符號；建立 PREPAID/COLLECT 下拉選單杜絕錯字",
        "content_type": "dual_image",
        "image_key_1": "fig6_find_and_replace_dialog.png",
        "caption_1": "圖 3-4：尋找與取代 (Ctrl+H) 批次拔除單位",
        "image_key_2": "fig7_data_validation_dialog.png",
        "caption_2": "圖 3-5：資料驗證設定合法清單下拉選單",
        "tips": "【源頭治理】：與其事後花時間辛苦清洗，不如在資料輸入端直接設『資料驗證清單』，強制行政人員只能選標準值！"
    },
    # Slide 14: Ch3 Filter and Sort
    {
        "id": "slide-14",
        "chapter": "第三章：30筆實戰示範",
        "title": "3.6 功能選項 6：【自動篩選與排序】快速嗅探髒資料",
        "subtitle": "展開篩選清單聚合唯一值，大小寫紊亂、隱形空格與空值瞬間現形",
        "content_type": "image_focus",
        "image_key": "fig8_filter_and_sort_ui.png",
        "caption": "圖 3-6：【自動篩選】快捷鍵【Ctrl + Shift + L】嗅探髒資料選單 (點擊圖片可放大)",
        "tips": "【偵錯心法】：按下 Ctrl+Shift+L 展開下拉清單，若同時看到 'TWKHH'、'twkhh' 與 ' TWKHH '，立刻知道該欄被污染了！"
    },
    # Slide 15: Ch3 Text Formula Chain
    {
        "id": "slide-15",
        "chapter": "第三章：30筆實戰示範",
        "title": "3.7 函數公式 1：文字清洗五大天王實戰鏈",
        "subtitle": "TRIM 清除首尾空格、UPPER 強制轉大寫、SUBSTITUTE 拔除單位、MID/FIND 動態截取",
        "content_type": "interactive_formula_playground",
        "image_key": "fig9_text_formulas_workflow.png",
        "caption": "圖 3-7：核心公式圖解：文字清洗五大天王函數實戰鏈",
        "tips": "【即時互動】：您可在下方輸入框自行輸入測試字串，即時體驗公式清洗結果！"
    },
    # Slide 16: Ch3 Numeric & Date Formulas
    {
        "id": "slide-16",
        "chapter": "第三章：30筆實戰示範",
        "title": "3.8 函數公式 2：數值淨化與日期時效計算 (VALUE & DATEVALUE)",
        "subtitle": "巢狀清除貨幣符號、統一混亂日期，直接以 ETA - ETD 計算航程天數",
        "content_type": "image_focus",
        "image_key": "fig10_date_functions_and_transit.png",
        "caption": "圖 3-8：日期前處理與航運時效運算三階段流程圖 (點擊圖片可放大)",
        "tips": "【實務公式】：=VALUE(SUBSTITUTE(I2, \"$\", \"\")) 可將文字型金額轉為實體數值；標準日期直接以 =ETA - ETD 即可算出精確航程天數。"
    },
    # Slide 17: Ch4 300-Row Architecture
    {
        "id": "slide-17",
        "chapter": "第四章：300筆大數據實戰",
        "title": "4.1 300 筆大數據航運實務案例情境全景架構",
        "subtitle": "涵蓋遠洋美東美西、歐洲地中海與亞洲航線，模擬大規模綜合清洗管線",
        "content_type": "image_focus",
        "image_key": "fig12_practice_300_workflow.png",
        "caption": "圖 4-1：航運大數據清洗專案架構地圖 (300 筆綜合管線) (點擊圖片可放大)",
        "tips": "【專案思維】：真實企業資料處理不是零散的，而是一條串聯『導入 -> 格式清洗 -> 數值化 -> 碼表關聯 -> 業務運算 -> 稽核』的管線。"
    },
    # Slide 18: Ch4 Master Tables & XLOOKUP
    {
        "id": "slide-18",
        "chapter": "第四章：300筆大數據實戰",
        "title": "4.2 航運核心碼表串接：XLOOKUP / VLOOKUP 動態查表",
        "subtitle": "港口 UN/LOCODE 關聯中文名、ISO 6346 櫃型代碼換算標準 TEU 係數",
        "content_type": "image_focus",
        "image_key": "fig11_xlookup_port_master.png",
        "caption": "圖 4-2：現代動態查表關聯：XLOOKUP / VLOOKUP 跨表串接港口基本檔 (點擊圖片可放大)",
        "tips": "【新舊對照】：Office 365 / Excel 2021+ 強烈推薦 =XLOOKUP，不再需要數第幾欄，且內建查無資料預設值，徹底告別 #N/A 錯誤！"
    },
    # Slide 19: Ch4 Tasks 1-4 Solutions
    {
        "id": "slide-19",
        "chapter": "第四章：300筆大數據實戰",
        "title": "4.3 任務挑戰精解 (一)：任務 1 至任務 4 標準解法卡",
        "subtitle": "提單標準化、港口名稱跨表關聯、貨櫃 TEU 換算與毛重數值化",
        "content_type": "task_cards_1_4",
        "tips": "【演講者備忘】：毛重欄位含有 kgs、kg 與逗號，公式採用三重 SUBSTITUTE 徹底淨化，最後由 VALUE 轉回數值。"
    },
    # Slide 20: Ch4 Tasks 5-8 Solutions
    {
        "id": "slide-20",
        "chapter": "第四章：300筆大數據實戰",
        "title": "4.4 任務挑戰精解 (二)：任務 5 至任務 8 標準解法卡",
        "subtitle": "運費試算、日期航期、港口免租期 Demurrage 滯港罰款預警與重複稽核",
        "content_type": "task_cards_5_8",
        "tips": "【業務亮點】：任務 7 滯港費預警利用 =MAX(0, (GateOut-ATA)-FreeDays) 搭配 IF 函數，一秒抓出需要向貨主追討滯港費的異常清單！"
    },
    # Slide 21: Demurrage Interactive Calculator
    {
        "id": "slide-21",
        "chapter": "第四章：300筆大數據實戰",
        "title": "4.5 航運實務試算器：滯港費 (Demurrage) 與航期即時運算機",
        "subtitle": "在線動態調整實際到港日、提櫃日與免租天數，即時檢視預警運算邏輯",
        "content_type": "interactive_demurrage_calc",
        "tips": "【演講者提示】：請學員嘗試輸入不同到港日與提櫃日，體驗 Excel 邏輯公式如何在幾毫秒內完成超期滯港判定！"
    },
    # Slide 22: Ch5 References & Cheat Sheet
    {
        "id": "slide-22",
        "chapter": "第五章：參考文獻與速查寶典",
        "title": "5.1 國際海運與資料治理參考文獻 & 王牌函數速查手冊",
        "subtitle": "遵循 ISO 6346、UN/LOCODE、DAMA-DMBOK2 與 Microsoft 官方指南",
        "content_type": "references_and_cheat_sheet",
        "tips": "【持續查閱】：本頁彙總了十大王牌前處理函數的標準語法與實務情境，建議同仁截圖或列印常備於辦公桌旁。"
    },
    # Slide 23: Conclusion & Next Steps
    {
        "id": "slide-23",
        "chapter": "結語與學習藍圖",
        "title": "航運行政人員資料分析進階學習藍圖",
        "subtitle": "從 Excel 基礎前處理，邁向 Power Query 自動化清洗與 Power BI 商業智慧",
        "content_type": "conclusion",
        "tips": "【結語心法】：乾淨的資料，是所有精準商業決策的唯一起點。恭喜各位完成資料前處理實戰訓練！"
    }
]

# Build HTML Presentation Application
html_content = f"""<!DOCTYPE html>
<html lang="zh-TW" class="dark">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>航運行政人員 MS Excel 基礎資料前處理操作實戰指引 - PPT 互動簡報</title>
  <script src="https://www.gstatic.com/antigravity/web/dev/tailwindcss.min.js"></script>
  <style>
    /* Custom Presentation Styling */
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;700;900&family=JetBrains+Mono:wght@400;600&display=swap');
    
    body {{
      font-family: 'Noto Sans TC', 'Microsoft JhengHei', sans-serif;
      user-select: none;
      -webkit-user-select: none;
    }}
    code, pre, .font-mono {{
      font-family: 'JetBrains Mono', Consolas, monospace;
    }}
    .slide-transition {{
      transition: opacity 0.3s cubic-bezier(0.4, 0, 0.2, 1), transform 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }}
    .custom-scrollbar::-webkit-scrollbar {{
      width: 6px;
      height: 6px;
    }}
    .custom-scrollbar::-webkit-scrollbar-track {{
      background: rgba(15, 23, 42, 0.5);
    }}
    .custom-scrollbar::-webkit-scrollbar-thumb {{
      background: rgba(56, 189, 248, 0.3);
      border-radius: 3px;
    }}
    .custom-scrollbar::-webkit-scrollbar-thumb:hover {{
      background: rgba(56, 189, 248, 0.6);
    }}
    .glass-panel {{
      background: rgba(15, 23, 42, 0.75);
      backdrop-filter: blur(12px);
      -webkit-backdrop-filter: blur(12px);
      border: 1px solid rgba(255, 255, 255, 0.1);
    }}
    .glow-blue {{
      box-shadow: 0 0 25px -5px rgba(56, 189, 248, 0.25);
    }}
    .glow-gold {{
      box-shadow: 0 0 25px -5px rgba(245, 158, 11, 0.25);
    }}
  </style>
</head>
<body class="bg-[#070D18] text-slate-100 min-h-screen flex flex-col justify-between overflow-x-hidden antialiased">

  <!-- TOP PROGRESS BAR & HEADER -->
  <header class="sticky top-0 z-40 bg-[#0A1324]/90 backdrop-blur-md border-b border-slate-800/80">
    <!-- Progress Indicator Line -->
    <div class="w-full bg-slate-800 h-1">
      <div id="progress-bar" class="bg-gradient-to-r from-cyan-500 via-blue-500 to-indigo-500 h-1 w-0 transition-all duration-300"></div>
    </div>
    
    <div class="max-w-7xl mx-auto px-4 sm:px-6 h-14 flex items-center justify-between gap-4">
      <!-- Left: Title & Chapter -->
      <div class="flex items-center gap-3 min-w-0">
        <span class="inline-flex items-center justify-center p-1.5 rounded-lg bg-blue-600/20 text-cyan-400 border border-blue-500/30">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path></svg>
        </span>
        <div class="truncate">
          <h1 class="text-xs sm:text-sm font-bold text-white tracking-wide truncate">
            航運行政人員 MS Excel 基礎資料前處理操作實戰指引
          </h1>
          <p id="header-chapter" class="text-[11px] text-cyan-400 font-medium">第一章：介面導航</p>
        </div>
      </div>

      <!-- Right: Navigation Controls -->
      <div class="flex items-center gap-1.5 sm:gap-2 flex-shrink-0">
        <span id="slide-readout" class="font-mono text-xs text-slate-400 bg-slate-800/60 px-2.5 py-1 rounded-md border border-slate-700/50">
          01 / 23
        </span>

        <button onclick="prevSlide()" title="上一張 (鍵盤 ←)" class="p-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white border border-slate-700 transition">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 19l-7-7 7-7"></path></svg>
        </button>

        <button onclick="nextSlide()" title="下一張 (鍵盤 → 或 空白鍵)" class="p-1.5 rounded-lg bg-blue-600 hover:bg-blue-500 text-white font-medium border border-blue-400/30 transition shadow-sm">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5l7 7-7 7"></path></svg>
        </button>

        <button onclick="toggleIndexModal()" title="章節目錄索引 (M)" class="p-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-cyan-400 border border-slate-700 transition flex items-center gap-1 text-xs">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 6h16M4 12h16M4 18h7"></path></svg>
          <span class="hidden md:inline">目錄</span>
        </button>

        <button onclick="toggleSpeakerTips()" title="顯示/隱藏實務備忘 (T)" id="tips-btn" class="p-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-amber-400 hover:text-amber-300 border border-slate-700 transition flex items-center gap-1 text-xs">
          <svg class="w-4 h-4" fill="currentColor" viewBox="0 0 20 20"><path d="M11 3a1 1 0 10-2 0v1a1 1 0 102 0V3zM15.657 5.757a1 1 0 00-1.414-1.414l-.707.707a1 1 0 001.414 1.414l.707-.707zM18 10a1 1 0 01-1 1h-1a1 1 0 110-2h1a1 1 0 011 1zM5.05 6.464A1 1 0 106.464 5.05l-.707-.707a1 1 0 00-1.414 1.414l.707.707zM5 10a1 1 0 01-1 1H3a1 1 0 110-2h1a1 1 0 011 1zM8 16v-1h4v1a2 2 0 11-4 0zM12 14c.015-.34.208-.646.477-.859a4 4 0 10-4.954 0c.27.213.462.519.477.859h4z"></path></svg>
          <span class="hidden md:inline">備忘</span>
        </button>

        <button onclick="toggleFullScreen()" title="全螢幕簡報 (F)" class="p-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white border border-slate-700 transition">
          <svg id="fullscreen-icon" class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 8V4m0 0h4M4 4l5 5m11-1V4m0 0h-4m4 0l-5 5M4 16v4m0 0h4m-4 0l5-5m11 5l-5-5m5 5v-4m0 4h-4"></path></svg>
        </button>
      </div>
    </div>
  </header>

  <!-- MAIN SLIDE DISPLAY STAGE -->
  <main class="flex-1 flex items-center justify-center p-3 sm:p-6 md:p-8 max-w-7xl w-full mx-auto">
    <div id="slide-stage" class="w-full glass-panel rounded-2xl p-6 sm:p-8 md:p-10 shadow-2xl relative overflow-hidden min-h-[580px] flex flex-col justify-between">
      <!-- Dynamic slide content will be injected here -->
      <div id="slide-content" class="w-full h-full flex-1"></div>

      <!-- Bottom Speaker Tips Panel (Collapsible) -->
      <div id="speaker-tips-box" class="mt-6 pt-4 border-t border-slate-800/80 bg-slate-900/60 -mx-6 -mb-6 sm:-mx-8 sm:-mb-8 md:-mx-10 md:-mb-10 px-6 sm:px-8 md:px-10 py-3.5 flex items-start gap-3 text-xs sm:text-sm text-slate-300 rounded-b-2xl">
        <span class="flex-shrink-0 p-1 rounded bg-amber-500/20 text-amber-400 mt-0.5">
          <svg class="w-4 h-4" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7-4a1 1 0 11-2 0 1 1 0 012 0zM9 9a1 1 0 000 2v3a1 1 0 001 1h1a1 1 0 100-2v-3a1 1 0 00-1-1H9z" clip-rule="evenodd"></path></svg>
        </span>
        <div class="flex-1">
          <span class="font-bold text-amber-400 mr-2">航運實務講者備忘：</span>
          <span id="speaker-tips-text" class="text-slate-300"></span>
        </div>
      </div>
    </div>
  </main>

  <!-- FOOTER CONTROLS & SHORTCUT HINTS -->
  <footer class="bg-[#0A1324] border-t border-slate-800/80 py-2.5 px-4 text-center text-xs text-slate-400 flex flex-wrap items-center justify-between gap-2 max-w-7xl mx-auto w-full">
    <div class="flex items-center gap-3">
      <span class="text-slate-400 font-medium">快捷鍵指南:</span>
      <span class="bg-slate-800 px-1.5 py-0.5 rounded text-slate-300 border border-slate-700">← / →</span> 切換投影片
      <span class="bg-slate-800 px-1.5 py-0.5 rounded text-slate-300 border border-slate-700">空白鍵</span> 下一張
      <span class="bg-slate-800 px-1.5 py-0.5 rounded text-slate-300 border border-slate-700">F</span> 全螢幕
      <span class="bg-slate-800 px-1.5 py-0.5 rounded text-slate-300 border border-slate-700">M</span> 目錄
    </div>
    <div class="text-slate-400">
      航運行政培訓小組 © 2026. Microsoft 365 / Excel 2021+ 實戰專用
    </div>
  </footer>

  <!-- IMAGE LIGHTBOX MODAL (FOR HIGH-RES ZOOM) -->
  <div id="image-modal" class="fixed inset-0 z-50 bg-black/90 backdrop-blur-md hidden items-center justify-center p-4" onclick="closeImageModal()">
    <div class="relative max-w-5xl w-full max-h-[90vh] flex flex-col items-center" onclick="event.stopPropagation()">
      <button onclick="closeImageModal()" class="absolute -top-10 right-0 text-white hover:text-cyan-400 text-sm font-semibold flex items-center gap-1 bg-slate-800/80 px-3 py-1 rounded-full border border-slate-600 transition">
        <span>關閉 (Esc)</span> ✕
      </button>
      <img id="modal-img" src="" alt="放大檢視" class="max-w-full max-h-[80vh] object-contain rounded-xl border border-slate-700 shadow-2xl bg-white/5">
      <p id="modal-caption" class="text-sm text-cyan-300 font-medium mt-3 text-center"></p>
    </div>
  </div>

  <!-- SLIDE INDEX JUMP MODAL -->
  <div id="index-modal" class="fixed inset-0 z-50 bg-black/80 backdrop-blur-md hidden items-center justify-center p-4" onclick="toggleIndexModal()">
    <div class="bg-slate-900 border border-slate-700 max-w-4xl w-full rounded-2xl p-6 max-h-[85vh] flex flex-col shadow-2xl" onclick="event.stopPropagation()">
      <div class="flex items-center justify-between pb-4 border-b border-slate-800">
        <h3 class="text-lg font-bold text-white flex items-center gap-2">
          <span class="w-2.5 h-2.5 rounded-full bg-cyan-400"></span>
          投影片全覽目錄 (Slide Directory)
        </h3>
        <button onclick="toggleIndexModal()" class="text-slate-400 hover:text-white p-1 rounded-lg">✕</button>
      </div>
      <div id="index-grid" class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3 overflow-y-auto py-4 custom-scrollbar flex-1 pr-1">
        <!-- Injected via JavaScript -->
      </div>
      <div class="pt-3 border-t border-slate-800 text-xs text-slate-400 flex justify-between items-center">
        <span>共 23 張實戰教學投影片</span>
        <span>點擊任一張立即跳轉</span>
      </div>
    </div>
  </div>

  <!-- SCRIPT LOGIC -->
  <script>
    const b64Images = {json.dumps(b64_images)};
    const slides = {json.dumps(slides_data)};
    
    let currentSlide = 0;
    let showTips = true;

    // Render current slide
    function renderSlide(index) {{
      if (index < 0) index = 0;
      if (index >= slides.length) index = slides.length - 1;
      currentSlide = index;

      const slide = slides[currentSlide];
      const stage = document.getElementById("slide-content");
      
      // Update header
      document.getElementById("header-chapter").innerText = slide.chapter;
      document.getElementById("slide-readout").innerText = `${{String(currentSlide + 1).padStart(2, '0')}} / ${{String(slides.length).padStart(2, '0')}}`;
      
      // Update progress bar
      const pct = ((currentSlide + 1) / slides.length) * 100;
      document.getElementById("progress-bar").style.width = pct + "%";

      // Update speaker tips
      document.getElementById("speaker-tips-text").innerText = slide.tips || "無特定備忘。";

      // Render slide content based on content_type
      let html = "";
      
      if (slide.content_type === "cover") {{
        html = `
          <div class="flex flex-col items-center justify-center text-center py-8 md:py-14 space-y-6">
            <span class="px-3.5 py-1.5 rounded-full bg-cyan-500/10 border border-cyan-500/30 text-cyan-400 text-xs font-semibold tracking-wider uppercase">
              ${{slide.badge}}
            </span>
            <h2 class="text-3xl sm:text-4xl md:text-5xl font-black text-white tracking-tight leading-tight max-w-4xl">
              ${{slide.title}}
            </h2>
            <div class="w-24 h-1 bg-gradient-to-r from-blue-500 to-cyan-400 rounded-full my-2"></div>
            <p class="text-slate-300 text-sm sm:text-base md:text-lg max-w-2xl font-normal leading-relaxed">
              ${{slide.subtitle}}
            </p>
            
            <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 pt-6 w-full max-w-3xl">
              <div class="bg-slate-800/60 border border-slate-700/60 p-3.5 rounded-xl text-center">
                <span class="text-xs text-slate-400 block">授課工具</span>
                <span class="text-sm font-bold text-white mt-1 block">MS Excel 365 / 2021+</span>
              </div>
              <div class="bg-slate-800/60 border border-slate-700/60 p-3.5 rounded-xl text-center">
                <span class="text-xs text-slate-400 block">教學受眾</span>
                <span class="text-sm font-bold text-cyan-400 mt-1 block">航運/海運行政人員</span>
              </div>
              <div class="bg-slate-800/60 border border-slate-700/60 p-3.5 rounded-xl text-center">
                <span class="text-xs text-slate-400 block">實作資料檔</span>
                <span class="text-sm font-bold text-amber-400 mt-1 block">30筆示範 + 300筆大數據</span>
              </div>
              <div class="bg-slate-800/60 border border-slate-700/60 p-3.5 rounded-xl text-center">
                <span class="text-xs text-slate-400 block">核心重點</span>
                <span class="text-sm font-bold text-emerald-400 mt-1 block">資料衛生與精準清洗</span>
              </div>
            </div>

            <div class="pt-4">
              <button onclick="nextSlide()" class="px-6 py-2.5 rounded-xl bg-gradient-to-r from-blue-600 to-cyan-600 hover:from-blue-500 hover:to-cyan-500 text-white font-bold text-sm shadow-lg transition flex items-center gap-2 mx-auto">
                開始探索課程 ➔
              </button>
            </div>
          </div>
        `;
      }} else if (slide.content_type === "agenda") {{
        html = `
          <div class="space-y-6">
            <div>
              <span class="text-xs font-bold text-cyan-400 uppercase tracking-wider">${{slide.chapter}}</span>
              <h2 class="text-2xl sm:text-3xl font-extrabold text-white mt-1">${{slide.title}}</h2>
              <p class="text-slate-400 text-xs sm:text-sm mt-1">${{slide.subtitle}}</p>
            </div>
            
            <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4 pt-2">
              <div onclick="goToSlide(2)" class="cursor-pointer bg-slate-800/50 hover:bg-slate-800 border border-slate-700/60 hover:border-cyan-500/50 p-4 rounded-xl transition group">
                <div class="flex items-center gap-2 mb-2">
                  <span class="w-6 h-6 rounded-md bg-blue-500/20 text-cyan-400 font-bold text-xs flex items-center justify-center">01</span>
                  <h3 class="font-bold text-white group-hover:text-cyan-300 text-sm">第一章：Excel 操作介面</h3>
                </div>
                <p class="text-xs text-slate-400">認識功能區、編輯列、名稱方塊、狀態列與十大必背快捷鍵。</p>
              </div>

              <div onclick="goToSlide(5)" class="cursor-pointer bg-slate-800/50 hover:bg-slate-800 border border-slate-700/60 hover:border-cyan-500/50 p-4 rounded-xl transition group">
                <div class="flex items-center gap-2 mb-2">
                  <span class="w-6 h-6 rounded-md bg-cyan-500/20 text-cyan-400 font-bold text-xs flex items-center justify-center">02</span>
                  <h3 class="font-bold text-white group-hover:text-cyan-300 text-sm">第二章：資料衛生常識</h3>
                </div>
                <p class="text-xs text-slate-400">GIGO原則、第一正規化(1NF)、航運七大骯髒資料病症與四大型態。</p>
              </div>

              <div onclick="goToSlide(8)" class="cursor-pointer bg-slate-800/50 hover:bg-slate-800 border border-slate-700/60 hover:border-cyan-500/50 p-4 rounded-xl transition group">
                <div class="flex items-center gap-2 mb-2">
                  <span class="w-6 h-6 rounded-md bg-emerald-500/20 text-emerald-400 font-bold text-xs flex items-center justify-center">03</span>
                  <h3 class="font-bold text-white group-hover:text-emerald-300 text-sm">第三章：30筆實戰示範</h3>
                </div>
                <p class="text-xs text-slate-400">資料剖析、Ctrl+E、移除重複、取代、資料驗證與文字清洗公式鏈。</p>
              </div>

              <div onclick="goToSlide(16)" class="cursor-pointer bg-slate-800/50 hover:bg-slate-800 border border-slate-700/60 hover:border-cyan-500/50 p-4 rounded-xl transition group">
                <div class="flex items-center gap-2 mb-2">
                  <span class="w-6 h-6 rounded-md bg-amber-500/20 text-amber-400 font-bold text-xs flex items-center justify-center">04</span>
                  <h3 class="font-bold text-white group-hover:text-amber-300 text-sm">第四章：300筆大數據全解</h3>
                </div>
                <p class="text-xs text-slate-400">UN/LOCODE、ISO TEU換算、八大任務逐題全解與滯港費預警運算。</p>
              </div>

              <div onclick="goToSlide(20)" class="cursor-pointer bg-slate-800/50 hover:bg-slate-800 border border-slate-700/60 hover:border-cyan-500/50 p-4 rounded-xl transition group">
                <div class="flex items-center gap-2 mb-2">
                  <span class="w-6 h-6 rounded-md bg-indigo-500/20 text-indigo-400 font-bold text-xs flex items-center justify-center">★</span>
                  <h3 class="font-bold text-white group-hover:text-indigo-300 text-sm">互動模擬運算機</h3>
                </div>
                <p class="text-xs text-slate-400">滯港費 (Demurrage) 與航期即時運算機，即時調整天數感受邏輯。</p>
              </div>

              <div onclick="goToSlide(21)" class="cursor-pointer bg-slate-800/50 hover:bg-slate-800 border border-slate-700/60 hover:border-cyan-500/50 p-4 rounded-xl transition group">
                <div class="flex items-center gap-2 mb-2">
                  <span class="w-6 h-6 rounded-md bg-purple-500/20 text-purple-400 font-bold text-xs flex items-center justify-center">05</span>
                  <h3 class="font-bold text-white group-hover:text-purple-300 text-sm">第五章：參考文獻速查表</h3>
                </div>
                <p class="text-xs text-slate-400">國際海運標準、DAMA資料治理規範、十大王牌函數速查表與學習藍圖。</p>
              </div>
            </div>
          </div>
        `;
      }} else if (slide.content_type === "why_prep") {{
        html = `
          <div class="space-y-6">
            <div>
              <span class="text-xs font-bold text-cyan-400 uppercase tracking-wider">${{slide.chapter}}</span>
              <h2 class="text-2xl sm:text-3xl font-extrabold text-white mt-1">${{slide.title}}</h2>
              <p class="text-slate-400 text-xs sm:text-sm mt-1">${{slide.subtitle}}</p>
            </div>

            <div class="grid grid-cols-1 md:grid-cols-2 gap-6 pt-2">
              <div class="bg-red-500/10 border border-red-500/30 p-5 rounded-xl space-y-3">
                <div class="flex items-center gap-2 text-red-400 font-bold text-base">
                  <span>❌</span> 未經前處理的航運災難實況
                </div>
                <ul class="space-y-2 text-xs sm:text-sm text-slate-300">
                  <li class="flex items-start gap-2">
                    <span class="text-red-400 mt-0.5">•</span>
                    <span><strong>肉眼看不出的空格：</strong>提單號碼多按了一格空白，VLOOKUP 永遠回傳 #N/A，造成報關延誤。</span>
                  </li>
                  <li class="flex items-start gap-2">
                    <span class="text-red-400 mt-0.5">•</span>
                    <span><strong>單位塞在數字中：</strong>貨物毛重手動打入 '18500 kg'，導致 SUM 加總為 0，運費請款算錯。</span>
                  </li>
                  <li class="flex items-start gap-2">
                    <span class="text-red-400 mt-0.5">•</span>
                    <span><strong>屬性混同在同一格：</strong>船名與航次擠在一起，主管要求按『單一船舶』分析運載量時徹底卡死。</span>
                  </li>
                  <li class="flex items-start gap-2">
                    <span class="text-red-400 mt-0.5">•</span>
                    <span><strong>重複提單建檔：</strong>同一票貨重複統計，出口運費營收虛胖，產生嚴重財務核算落差。</span>
                  </li>
                </ul>
              </div>

              <div class="bg-emerald-500/10 border border-emerald-500/30 p-5 rounded-xl space-y-3">
                <div class="flex items-center gap-2 text-emerald-400 font-bold text-base">
                  <span>✅</span> 掌握標準前處理的效益飛躍
                </div>
                <ul class="space-y-2 text-xs sm:text-sm text-slate-300">
                  <li class="flex items-start gap-2">
                    <span class="text-emerald-400 mt-0.5">•</span>
                    <span><strong>工時省下 80%：</strong>從過去人工手動核對數百筆，轉變為幾秒鐘批次自動化清洗完成。</span>
                  </li>
                  <li class="flex items-start gap-2">
                    <span class="text-emerald-400 mt-0.5">•</span>
                    <span><strong>樞紐分析一鍵產出：</strong>標準一維二維表結構，讓交叉分析篩選器與圖表隨點隨動。</span>
                  </li>
                  <li class="flex items-start gap-2">
                    <span class="text-emerald-400 mt-0.5">•</span>
                    <span><strong>跨系統無縫對接：</strong>清洗後的 UN/LOCODE 與 ISO 貨櫃編碼，與各大海關 EDI 系統 100% 相容。</span>
                  </li>
                  <li class="flex items-start gap-2">
                    <span class="text-emerald-400 mt-0.5">•</span>
                    <span><strong>精確預警滯港費用：</strong>到港日與提櫃日自動比對免租期，一鍵過濾出 Demurrage 罰款名單。</span>
                  </li>
                </ul>
              </div>
            </div>

            <div class="p-4 rounded-xl bg-slate-800/70 border border-slate-700 text-xs sm:text-sm text-slate-300 flex items-center gap-3">
              <span class="text-cyan-400 text-lg font-bold">💡 核心金句：</span>
              <span>『Garbage In, Garbage Out』—— 缺乏品質前處理的數據，算出來的報表再漂亮，本質上依舊是錯誤決策的垃圾！</span>
            </div>
          </div>
        `;
      }} else if (slide.content_type === "shortcuts") {{
        html = `
          <div class="space-y-5">
            <div>
              <span class="text-xs font-bold text-cyan-400 uppercase tracking-wider">${{slide.chapter}}</span>
              <h2 class="text-2xl sm:text-3xl font-extrabold text-white mt-1">${{slide.title}}</h2>
              <p class="text-slate-400 text-xs sm:text-sm mt-1">${{slide.subtitle}}</p>
            </div>

            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3.5 pt-1">
              <div class="bg-slate-800/60 border border-slate-700 p-3.5 rounded-xl flex items-start gap-3">
                <span class="font-mono text-xs font-bold text-amber-300 bg-amber-500/20 px-2 py-1 rounded border border-amber-500/30 whitespace-nowrap">Ctrl+Shift+L</span>
                <div>
                  <h4 class="text-sm font-bold text-white">啟動/關閉自動篩選</h4>
                  <p class="text-xs text-slate-400 mt-0.5">一秒為所有標題加上下拉箭頭，嗅探髒資料唯一值首選。</p>
                </div>
              </div>

              <div class="bg-slate-800/60 border border-slate-700 p-3.5 rounded-xl flex items-start gap-3">
                <span class="font-mono text-xs font-bold text-cyan-300 bg-cyan-500/20 px-2 py-1 rounded border border-cyan-500/30 whitespace-nowrap">Ctrl + E</span>
                <div>
                  <h4 class="text-sm font-bold text-white">智慧快速填入 (Flash Fill)</h4>
                  <p class="text-xs text-slate-400 mt-0.5">輸入第一筆範本後按下，Excel 自動嗅探規律秒速填滿全欄！</p>
                </div>
              </div>

              <div class="bg-slate-800/60 border border-slate-700 p-3.5 rounded-xl flex items-start gap-3">
                <span class="font-mono text-xs font-bold text-blue-300 bg-blue-500/20 px-2 py-1 rounded border border-blue-500/30 whitespace-nowrap">Ctrl + H</span>
                <div>
                  <h4 class="text-sm font-bold text-white">尋找與取代</h4>
                  <p class="text-xs text-slate-400 mt-0.5">批次將 'kg'、'$' 取代為空白，瞬間純數值化。</p>
                </div>
              </div>

              <div class="bg-slate-800/60 border border-slate-700 p-3.5 rounded-xl flex items-start gap-3">
                <span class="font-mono text-xs font-bold text-emerald-300 bg-emerald-500/20 px-2 py-1 rounded border border-emerald-500/30 whitespace-nowrap">Ctrl + 1</span>
                <div>
                  <h4 class="text-sm font-bold text-white">儲存格格式設定</h4>
                  <p class="text-xs text-slate-400 mt-0.5">自訂日期 YYYY-MM-DD，千分位逗號與小數位數。</p>
                </div>
              </div>

              <div class="bg-slate-800/60 border border-slate-700 p-3.5 rounded-xl flex items-start gap-3">
                <span class="font-mono text-xs font-bold text-purple-300 bg-purple-500/20 px-2 py-1 rounded border border-purple-500/30 whitespace-nowrap">Ctrl + 方向鍵</span>
                <div>
                  <h4 class="text-sm font-bold text-white">飛躍至資料邊界</h4>
                  <p class="text-xs text-slate-400 mt-0.5">在數千筆提單中，一秒跳轉至最頂端或最後一筆資料。</p>
                </div>
              </div>

              <div class="bg-slate-800/60 border border-slate-700 p-3.5 rounded-xl flex items-start gap-3">
                <span class="font-mono text-xs font-bold text-rose-300 bg-rose-500/20 px-2 py-1 rounded border border-rose-500/30 whitespace-nowrap">Alt + = (等號)</span>
                <div>
                  <h4 class="text-sm font-bold text-white">自動加總 (AutoSum)</h4>
                  <p class="text-xs text-slate-400 mt-0.5">在運費或重量底端，一鍵自動產生 SUM 加總公式。</p>
                </div>
              </div>
            </div>
          </div>
        `;
      }} else if (slide.content_type === "image_focus") {{
        const b64 = b64Images[slide.image_key] || "";
        html = `
          <div class="space-y-4">
            <div>
              <span class="text-xs font-bold text-cyan-400 uppercase tracking-wider">${{slide.chapter}}</span>
              <h2 class="text-xl sm:text-2xl font-extrabold text-white mt-1">${{slide.title}}</h2>
              <p class="text-slate-400 text-xs sm:text-sm mt-0.5">${{slide.subtitle}}</p>
            </div>

            <div class="flex flex-col items-center">
              <div class="relative group cursor-pointer max-w-4xl w-full bg-slate-900/80 p-2 rounded-xl border border-slate-700 shadow-xl overflow-hidden" onclick="openImageModal('${{b64}}', '${{slide.caption}}')">
                <img src="${{b64}}" alt="${{slide.title}}" class="w-full max-h-[380px] object-contain rounded-lg transition duration-200 group-hover:scale-[1.01]">
                <div class="absolute inset-0 bg-blue-600/10 opacity-0 group-hover:opacity-100 transition flex items-center justify-center">
                  <span class="px-3 py-1.5 rounded-lg bg-black/75 text-cyan-300 text-xs font-semibold backdrop-blur-sm border border-cyan-500/30">
                    🔍 點擊全螢幕放大檢視
                  </span>
                </div>
              </div>
              <p class="text-xs text-slate-400 mt-2 font-medium">${{slide.caption}}</p>
            </div>
          </div>
        `;
      }} else if (slide.content_type === "dual_image") {{
        const b64_1 = b64Images[slide.image_key_1] || "";
        const b64_2 = b64Images[slide.image_key_2] || "";
        html = `
          <div class="space-y-4">
            <div>
              <span class="text-xs font-bold text-cyan-400 uppercase tracking-wider">${{slide.chapter}}</span>
              <h2 class="text-xl sm:text-2xl font-extrabold text-white mt-1">${{slide.title}}</h2>
              <p class="text-slate-400 text-xs sm:text-sm mt-0.5">${{slide.subtitle}}</p>
            </div>

            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div class="flex flex-col items-center">
                <div class="relative group cursor-pointer w-full bg-slate-900/80 p-2 rounded-xl border border-slate-700 shadow-xl overflow-hidden" onclick="openImageModal('${{b64_1}}', '${{slide.caption_1}}')">
                  <img src="${{b64_1}}" alt="Image 1" class="w-full max-h-[300px] object-contain rounded-lg">
                  <div class="absolute inset-0 bg-blue-600/10 opacity-0 group-hover:opacity-100 transition flex items-center justify-center">
                    <span class="px-2.5 py-1 rounded bg-black/75 text-cyan-300 text-xs font-semibold">🔍 放大檢視</span>
                  </div>
                </div>
                <p class="text-xs text-slate-400 mt-2 font-medium">${{slide.caption_1}}</p>
              </div>

              <div class="flex flex-col items-center">
                <div class="relative group cursor-pointer w-full bg-slate-900/80 p-2 rounded-xl border border-slate-700 shadow-xl overflow-hidden" onclick="openImageModal('${{b64_2}}', '${{slide.caption_2}}')">
                  <img src="${{b64_2}}" alt="Image 2" class="w-full max-h-[300px] object-contain rounded-lg">
                  <div class="absolute inset-0 bg-blue-600/10 opacity-0 group-hover:opacity-100 transition flex items-center justify-center">
                    <span class="px-2.5 py-1 rounded bg-black/75 text-cyan-300 text-xs font-semibold">🔍 放大檢視</span>
                  </div>
                </div>
                <p class="text-xs text-slate-400 mt-2 font-medium">${{slide.caption_2}}</p>
              </div>
            </div>
          </div>
        `;
      }} else if (slide.content_type === "interactive_dirty_data") {{
        html = `
          <div class="space-y-4">
            <div>
              <span class="text-xs font-bold text-cyan-400 uppercase tracking-wider">${{slide.chapter}}</span>
              <h2 class="text-2xl font-extrabold text-white mt-1">${{slide.title}}</h2>
              <p class="text-slate-400 text-xs mt-0.5">${{slide.subtitle}}</p>
            </div>

            <!-- Tabs / Buttons -->
            <div class="flex flex-wrap gap-2 pt-1" id="dirty-tabs">
              <button onclick="showDirtyDetail(0)" class="dirty-tab-btn active px-3 py-1.5 rounded-lg text-xs font-bold transition bg-cyan-600 text-white">1. 不可見空格</button>
              <button onclick="showDirtyDetail(1)" class="dirty-tab-btn px-3 py-1.5 rounded-lg text-xs font-bold transition bg-slate-800 text-slate-300 hover:bg-slate-700">2. 大小寫混亂</button>
              <button onclick="showDirtyDetail(2)" class="dirty-tab-btn px-3 py-1.5 rounded-lg text-xs font-bold transition bg-slate-800 text-slate-300 hover:bg-slate-700">3. 單欄多屬性</button>
              <button onclick="showDirtyDetail(3)" class="dirty-tab-btn px-3 py-1.5 rounded-lg text-xs font-bold transition bg-slate-800 text-slate-300 hover:bg-slate-700">4. 數字存文字</button>
              <button onclick="showDirtyDetail(4)" class="dirty-tab-btn px-3 py-1.5 rounded-lg text-xs font-bold transition bg-slate-800 text-slate-300 hover:bg-slate-700">5. 混亂日期</button>
              <button onclick="showDirtyDetail(5)" class="dirty-tab-btn px-3 py-1.5 rounded-lg text-xs font-bold transition bg-slate-800 text-slate-300 hover:bg-slate-700">6. 重複建單</button>
              <button onclick="showDirtyDetail(6)" class="dirty-tab-btn px-3 py-1.5 rounded-lg text-xs font-bold transition bg-slate-800 text-slate-300 hover:bg-slate-700">7. 空白遺漏值</button>
            </div>

            <!-- Detail Display Panel -->
            <div id="dirty-detail-box" class="bg-slate-800/80 border border-slate-700 p-5 rounded-xl grid grid-cols-1 md:grid-cols-3 gap-4">
              <!-- Content updated via JS -->
            </div>
          </div>
        `;
      }} else if (slide.content_type === "interactive_table_toggle") {{
        html = `
          <div class="space-y-4">
            <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div>
                <span class="text-xs font-bold text-cyan-400 uppercase tracking-wider">${{slide.chapter}}</span>
                <h2 class="text-2xl font-extrabold text-white mt-1">${{slide.title}}</h2>
                <p class="text-slate-400 text-xs mt-0.5">${{slide.subtitle}}</p>
              </div>

              <!-- Toggle Button -->
              <div class="inline-flex p-1 rounded-xl bg-slate-800 border border-slate-700">
                <button id="btn-raw" onclick="switchTableMode('raw')" class="px-3 py-1.5 rounded-lg text-xs font-bold transition bg-red-600 text-white">
                  原始缺陷資料 (Raw)
                </button>
                <button id="btn-clean" onclick="switchTableMode('clean')" class="px-3 py-1.5 rounded-lg text-xs font-bold transition text-slate-400 hover:text-white">
                  清洗後標準資料 (Cleaned)
                </button>
              </div>
            </div>

            <!-- Table Container -->
            <div class="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden shadow-xl max-h-[360px] overflow-y-auto custom-scrollbar">
              <div id="table-display-area"></div>
            </div>
          </div>
        `;
      }} else if (slide.content_type === "interactive_formula_playground") {{
        const b64 = b64Images[slide.image_key] || "";
        html = `
          <div class="space-y-4">
            <div>
              <span class="text-xs font-bold text-cyan-400 uppercase tracking-wider">${{slide.chapter}}</span>
              <h2 class="text-2xl font-extrabold text-white mt-1">${{slide.title}}</h2>
              <p class="text-slate-400 text-xs mt-0.5">${{slide.subtitle}}</p>
            </div>

            <!-- Interactive Simulator Grid -->
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
              <!-- Left: Live Input Simulator -->
              <div class="bg-slate-800/70 border border-slate-700 p-4 rounded-xl space-y-3">
                <div class="flex items-center justify-between">
                  <h4 class="text-sm font-bold text-cyan-300">🧪 文字清洗即時試驗場</h4>
                  <span class="text-[11px] text-slate-400">可自由手動修改下方文字</span>
                </div>
                
                <div>
                  <label class="text-xs text-slate-400 block mb-1">測試輸入值 (含空格與斜線)：</label>
                  <input id="test-str-input" type="text" value="  EVER GIVEN / 0142-052W  " oninput="updateFormulaDemo()" 
                         class="w-full bg-slate-900 border border-cyan-500/50 rounded-lg px-3 py-2 text-sm font-mono text-cyan-300 focus:outline-none focus:ring-1 focus:ring-cyan-400">
                </div>

                <div class="space-y-2 text-xs font-mono pt-1">
                  <div class="p-2 rounded bg-slate-900 border border-slate-800">
                    <span class="text-slate-400 block text-[10px]">=TRIM(文字) [拔除首尾空格]</span>
                    <span id="res-trim" class="text-emerald-400 font-bold text-sm"></span>
                  </div>

                  <div class="p-2 rounded bg-slate-900 border border-slate-800">
                    <span class="text-slate-400 block text-[10px]">=TRIM(LEFT(C2, FIND("/", C2)-1)) [截取船名]</span>
                    <span id="res-vessel" class="text-cyan-300 font-bold text-sm"></span>
                  </div>

                  <div class="p-2 rounded bg-slate-900 border border-slate-800">
                    <span class="text-slate-400 block text-[10px]">=TRIM(MID(C2, FIND("/", C2)+1, 50)) [截取航次]</span>
                    <span id="res-voyage" class="text-amber-300 font-bold text-sm"></span>
                  </div>

                  <div class="p-2 rounded bg-slate-900 border border-slate-800">
                    <span class="text-slate-400 block text-[10px]">=UPPER(TRIM(文字)) [強制大寫]</span>
                    <span id="res-upper" class="text-purple-300 font-bold text-sm"></span>
                  </div>
                </div>
              </div>

              <!-- Right: Diagram Thumbnail -->
              <div class="flex flex-col items-center justify-center bg-slate-900/60 border border-slate-800 p-3 rounded-xl">
                <div class="relative group cursor-pointer max-w-full" onclick="openImageModal('${{b64}}', '${{slide.caption}}')">
                  <img src="${{b64}}" alt="Formula Diagram" class="w-full max-h-[300px] object-contain rounded-lg">
                  <div class="absolute inset-0 bg-blue-600/10 opacity-0 group-hover:opacity-100 transition flex items-center justify-center">
                    <span class="px-2.5 py-1 rounded bg-black/75 text-cyan-300 text-xs font-semibold">🔍 放大檢視原圖</span>
                  </div>
                </div>
                <p class="text-xs text-slate-400 mt-2">${{slide.caption}}</p>
              </div>
            </div>
          </div>
        `;
      }} else if (slide.content_type === "task_cards_1_4") {{
        html = `
          <div class="space-y-4">
            <div>
              <span class="text-xs font-bold text-cyan-400 uppercase tracking-wider">${{slide.chapter}}</span>
              <h2 class="text-2xl font-extrabold text-white mt-1">${{slide.title}}</h2>
              <p class="text-slate-400 text-xs mt-0.5">${{slide.subtitle}}</p>
            </div>

            <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div class="bg-slate-800/80 border border-slate-700 p-4 rounded-xl space-y-2">
                <span class="px-2 py-0.5 rounded bg-blue-500/20 text-cyan-400 text-xs font-bold">任務一</span>
                <h4 class="text-sm font-bold text-white">提單與訂艙編號標準化</h4>
                <p class="text-xs text-slate-400">消除首尾多餘空格，英文字母全數轉為大寫標準格式。</p>
                <div class="p-2 rounded bg-slate-900 border border-slate-800 font-mono text-xs text-cyan-300 font-bold">
                  =TRIM(UPPER(C2))
                </div>
              </div>

              <div class="bg-slate-800/80 border border-slate-700 p-4 rounded-xl space-y-2">
                <span class="px-2 py-0.5 rounded bg-cyan-500/20 text-cyan-400 text-xs font-bold">任務二</span>
                <h4 class="text-sm font-bold text-white">裝卸港代碼與中文名跨表關聯</h4>
                <p class="text-xs text-slate-400">POL/POD代碼大寫清洗，利用 XLOOKUP 跨表自動帶出港口中文名。</p>
                <div class="p-2 rounded bg-slate-900 border border-slate-800 font-mono text-xs text-emerald-400 font-bold">
                  =XLOOKUP(UPPER(TRIM(F2)), Port_Master!$A$2:$A$18, Port_Master!$B$2:$B$18, "未知港口")
                </div>
              </div>

              <div class="bg-slate-800/80 border border-slate-700 p-4 rounded-xl space-y-2">
                <span class="px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-400 text-xs font-bold">任務三</span>
                <h4 class="text-sm font-bold text-white">貨櫃號碼與 ISO 櫃型換算 TEU</h4>
                <p class="text-xs text-slate-400">依據 20GP=1, 40GP/40HC=2, 45HC=2.25 係數跨表換算運載量。</p>
                <div class="p-2 rounded bg-slate-900 border border-slate-800 font-mono text-xs text-amber-300 font-bold">
                  =XLOOKUP(I2, ISO_Type!$A$2:$A$9, ISO_Type!$E$2:$E$9, 1.0)
                </div>
              </div>

              <div class="bg-slate-800/80 border border-slate-700 p-4 rounded-xl space-y-2">
                <span class="px-2 py-0.5 rounded bg-purple-500/20 text-purple-400 text-xs font-bold">任務四</span>
                <h4 class="text-sm font-bold text-white">貨物毛重與包裝單位純數值化</h4>
                <p class="text-xs text-slate-400">剔除 'kg'、'KGS' 與逗號，將文字型數字轉回標準可加總數值。</p>
                <div class="p-2 rounded bg-slate-900 border border-slate-800 font-mono text-xs text-purple-300 font-bold">
                  =VALUE(SUBSTITUTE(SUBSTITUTE(LOWER(M2), "kg", ""), ",", ""))
                </div>
              </div>
            </div>
          </div>
        `;
      }} else if (slide.content_type === "task_cards_5_8") {{
        html = `
          <div class="space-y-4">
            <div>
              <span class="text-xs font-bold text-cyan-400 uppercase tracking-wider">${{slide.chapter}}</span>
              <h2 class="text-2xl font-extrabold text-white mt-1">${{slide.title}}</h2>
              <p class="text-slate-400 text-xs mt-0.5">${{slide.subtitle}}</p>
            </div>

            <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div class="bg-slate-800/80 border border-slate-700 p-4 rounded-xl space-y-2">
                <span class="px-2 py-0.5 rounded bg-blue-500/20 text-cyan-400 text-xs font-bold">任務五</span>
                <h4 class="text-sm font-bold text-white">海運費淨化與總運費試算</h4>
                <p class="text-xs text-slate-400">拔除 '$' 與 'USD' 幣別代碼，將淨化運費與附加費相加計算。</p>
                <div class="p-2 rounded bg-slate-900 border border-slate-800 font-mono text-xs text-cyan-300 font-bold">
                  =VALUE(SUBSTITUTE(SUBSTITUTE(P2, "$", ""), "USD", "")) + Q2
                </div>
              </div>

              <div class="bg-slate-800/80 border border-slate-700 p-4 rounded-xl space-y-2">
                <span class="px-2 py-0.5 rounded bg-cyan-500/20 text-cyan-400 text-xs font-bold">任務六</span>
                <h4 class="text-sm font-bold text-white">多樣化日期清洗與航期天數計算</h4>
                <p class="text-xs text-slate-400">文字型日期轉為序列值，直接相減得出正確航期天數。</p>
                <div class="p-2 rounded bg-slate-900 border border-slate-800 font-mono text-xs text-emerald-400 font-bold">
                  航程天數公式: =標準ETA - 標準ETD  (如 2026/10/22 - 2026/10/05 = 17天)
                </div>
              </div>

              <div class="bg-slate-800/80 border border-slate-700 p-4 rounded-xl space-y-2">
                <span class="px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-400 text-xs font-bold">任務七</span>
                <h4 class="text-sm font-bold text-white">免租期與滯港費 (Demurrage) 預警</h4>
                <p class="text-xs text-slate-400">停留天數超過免租期者，使用 MAX 與 IF 函數自動亮紅燈預警。</p>
                <div class="p-2 rounded bg-slate-900 border border-slate-800 font-mono text-xs text-amber-300 font-bold">
                  =IF(MAX(0, (GateOut - ATA) - FreeDays) > 0, "滯港罰款", "正常")
                </div>
              </div>

              <div class="bg-slate-800/80 border border-slate-700 p-4 rounded-xl space-y-2">
                <span class="px-2 py-0.5 rounded bg-purple-500/20 text-purple-400 text-xs font-bold">任務八</span>
                <h4 class="text-sm font-bold text-white">全域重複排查與缺漏資料稽核</h4>
                <p class="text-xs text-slate-400">排查重複提單，空客戶填 '[待確認客戶]'，空海關填 'PENDING'。</p>
                <div class="p-2 rounded bg-slate-900 border border-slate-800 font-mono text-xs text-purple-300 font-bold">
                  =IF(COUNTIF($C$2:C2, C2) > 1, "【重複建單】", "正常")
                </div>
              </div>
            </div>
          </div>
        `;
      }} else if (slide.content_type === "interactive_demurrage_calc") {{
        html = `
          <div class="space-y-4">
            <div>
              <span class="text-xs font-bold text-cyan-400 uppercase tracking-wider">${{slide.chapter}}</span>
              <h2 class="text-2xl font-extrabold text-white mt-1">${{slide.title}}</h2>
              <p class="text-slate-400 text-xs mt-0.5">${{slide.subtitle}}</p>
            </div>

            <!-- Interactive Calculator Box -->
            <div class="bg-slate-800/80 border border-slate-700 p-5 rounded-2xl grid grid-cols-1 md:grid-cols-2 gap-6">
              <!-- Inputs Column -->
              <div class="space-y-4">
                <h4 class="text-sm font-bold text-cyan-300 flex items-center gap-2">
                  <span>📅</span> 航運港口時效參數輸入
                </h4>

                <div>
                  <label class="text-xs text-slate-400 block mb-1">實際到港日 (ATA):</label>
                  <input id="calc-ata" type="date" value="2026-10-15" onchange="runDemurrageCalc()"
                         class="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-sm text-white font-mono focus:border-cyan-400 focus:outline-none">
                </div>

                <div>
                  <label class="text-xs text-slate-400 block mb-1">實際提櫃日 (Gate-Out):</label>
                  <input id="calc-gate" type="date" value="2026-10-28" onchange="runDemurrageCalc()"
                         class="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-sm text-white font-mono focus:border-cyan-400 focus:outline-none">
                </div>

                <div>
                  <label class="text-xs text-slate-400 block mb-1">船公司免租天數 (Free Days):</label>
                  <select id="calc-free" onchange="runDemurrageCalc()" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-sm text-white font-mono focus:border-cyan-400 focus:outline-none">
                    <option value="7">7 天 (一般乾貨櫃標準免租期)</option>
                    <option value="10">10 天 (大客戶特約合約期)</option>
                    <option value="14" selected>14 天 (跨國長約免租期)</option>
                    <option value="21">21 天 (特殊專案免租期)</option>
                  </select>
                </div>
              </div>

              <!-- Live Calculated Outputs Column -->
              <div class="bg-slate-900/80 border border-slate-800 p-4 rounded-xl flex flex-col justify-between space-y-3">
                <h4 class="text-sm font-bold text-slate-300">📊 運算結果即時判定看板</h4>

                <div class="grid grid-cols-2 gap-3">
                  <div class="p-3 rounded-lg bg-slate-800/60 border border-slate-700">
                    <span class="text-[11px] text-slate-400 block">港口停留天數</span>
                    <span id="out-stay" class="text-xl font-mono font-bold text-white mt-1 block">13 天</span>
                  </div>

                  <div class="p-3 rounded-lg bg-slate-800/60 border border-slate-700">
                    <span class="text-[11px] text-slate-400 block">超期天數 (Overdue)</span>
                    <span id="out-overdue" class="text-xl font-mono font-bold text-amber-400 mt-1 block">0 天</span>
                  </div>
                </div>

                <div id="demurrage-badge" class="p-4 rounded-xl text-center border font-bold text-base transition flex flex-col items-center justify-center gap-1">
                  <!-- Injected via JS -->
                </div>

                <p class="text-[11px] text-slate-400 font-mono">
                  底層公式：=IF(MAX(0, (Gate-ATA)-FreeDays)>0, "滯港罰款", "正常")
                </p>
              </div>
            </div>
          </div>
        `;
      }} else if (slide.content_type === "references_and_cheat_sheet") {{
        html = `
          <div class="space-y-4">
            <div>
              <span class="text-xs font-bold text-cyan-400 uppercase tracking-wider">${{slide.chapter}}</span>
              <h2 class="text-2xl font-extrabold text-white mt-1">${{slide.title}}</h2>
              <p class="text-slate-400 text-xs mt-0.5">${{slide.subtitle}}</p>
            </div>

            <!-- Table of cheat sheet -->
            <div class="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden max-h-[360px] overflow-y-auto custom-scrollbar">
              <table class="w-full text-left text-xs">
                <thead class="bg-slate-800 text-slate-300 font-bold sticky top-0">
                  <tr>
                    <th class="p-2.5">函數名稱</th>
                    <th class="p-2.5">標準語法結構</th>
                    <th class="p-2.5">功能分類</th>
                    <th class="p-2.5">航運行政實務應用情境</th>
                  </tr>
                </thead>
                <tbody class="divide-y divide-slate-800 text-slate-300">
                  <tr class="hover:bg-slate-800/50">
                    <td class="p-2.5 font-bold text-cyan-400 font-mono">TRIM</td>
                    <td class="p-2.5 font-mono text-slate-400">=TRIM(文字)</td>
                    <td class="p-2.5">文字清洗</td>
                    <td class="p-2.5">剔除提單號碼、貨櫃號碼與港口代碼的首尾隱形空格。</td>
                  </tr>
                  <tr class="hover:bg-slate-800/50">
                    <td class="p-2.5 font-bold text-cyan-400 font-mono">UPPER</td>
                    <td class="p-2.5 font-mono text-slate-400">=UPPER(文字)</td>
                    <td class="p-2.5">格式轉換</td>
                    <td class="p-2.5">將全小寫港口 'twkhh' 強制轉為標準大寫 'TWKHH'。</td>
                  </tr>
                  <tr class="hover:bg-slate-800/50">
                    <td class="p-2.5 font-bold text-cyan-400 font-mono">SUBSTITUTE</td>
                    <td class="p-2.5 font-mono text-slate-400">=SUBSTITUTE(文字, 舊字, 新字)</td>
                    <td class="p-2.5">字串替換</td>
                    <td class="p-2.5">拔除毛重中的 'kg' 單位，或將日期中的 '.' 替換為 '/'。</td>
                  </tr>
                  <tr class="hover:bg-slate-800/50">
                    <td class="p-2.5 font-bold text-cyan-400 font-mono">VALUE</td>
                    <td class="p-2.5 font-mono text-slate-400">=VALUE(文字數字)</td>
                    <td class="p-2.5">型態轉換</td>
                    <td class="p-2.5">將拔除幣別與單位後的文字型數字強制轉為可計算之數值。</td>
                  </tr>
                  <tr class="hover:bg-slate-800/50">
                    <td class="p-2.5 font-bold text-cyan-400 font-mono">DATEVALUE</td>
                    <td class="p-2.5 font-mono text-slate-400">=DATEVALUE(日期字串)</td>
                    <td class="p-2.5">日期轉換</td>
                    <td class="p-2.5">將多樣化格式之日期字串轉換為 Excel 內部標準日期序列值。</td>
                  </tr>
                  <tr class="hover:bg-slate-800/50">
                    <td class="p-2.5 font-bold text-cyan-400 font-mono">XLOOKUP</td>
                    <td class="p-2.5 font-mono text-slate-400">=XLOOKUP(目標, 查尋欄, 傳回欄)</td>
                    <td class="p-2.5">動態查表</td>
                    <td class="p-2.5">根據 5 碼 LOCODE 瞬間帶出港口中文名，免除 #N/A。</td>
                  </tr>
                  <tr class="hover:bg-slate-800/50">
                    <td class="p-2.5 font-bold text-cyan-400 font-mono">IF / OR</td>
                    <td class="p-2.5 font-mono text-slate-400">=IF(條件, 真值, 假值)</td>
                    <td class="p-2.5">邏輯防錯</td>
                    <td class="p-2.5">判斷託運人或付款狀態是否留白，自動填補預設審計標記。</td>
                  </tr>
                  <tr class="hover:bg-slate-800/50">
                    <td class="p-2.5 font-bold text-cyan-400 font-mono">COUNTIF</td>
                    <td class="p-2.5 font-mono text-slate-400">=COUNTIF(範圍, 條件)</td>
                    <td class="p-2.5">統計查重</td>
                    <td class="p-2.5">在連續範圍內偵測提單出現次數是否 > 1，抓出重複建單。</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        `;
      }} else if (slide.content_type === "conclusion") {{
        html = `
          <div class="space-y-6 py-2">
            <div>
              <span class="text-xs font-bold text-cyan-400 uppercase tracking-wider">${{slide.chapter}}</span>
              <h2 class="text-2xl sm:text-3xl font-extrabold text-white mt-1">${{slide.title}}</h2>
              <p class="text-slate-400 text-xs sm:text-sm mt-1">${{slide.subtitle}}</p>
            </div>

            <!-- 3 Stages Roadmap -->
            <div class="grid grid-cols-1 md:grid-cols-3 gap-4 pt-2">
              <div class="bg-slate-800/70 border border-slate-700 p-5 rounded-xl space-y-2">
                <span class="px-2 py-0.5 rounded bg-blue-500/20 text-cyan-400 text-xs font-bold">第一階段 (目前)</span>
                <h4 class="text-base font-bold text-white">奠定資料衛生基石</h4>
                <p class="text-xs text-slate-300">熟練一維二維表結構、快捷鍵、TRIM/UPPER/VALUE/XLOOKUP 函數鏈與源頭資料驗證。</p>
              </div>

              <div class="bg-slate-800/70 border border-slate-700 p-5 rounded-xl space-y-2">
                <span class="px-2 py-0.5 rounded bg-cyan-500/20 text-cyan-400 text-xs font-bold">第二階段 (進階)</span>
                <h4 class="text-base font-bold text-white">Power Query 自動化</h4>
                <p class="text-xs text-slate-300">學習 Excel 樞紐分析表、交叉分析篩選器，並運用 Power Query 建立『一鍵刷新清洗』流程。</p>
              </div>

              <div class="bg-slate-800/70 border border-slate-700 p-5 rounded-xl space-y-2">
                <span class="px-2 py-0.5 rounded bg-amber-500/20 text-amber-400 text-xs font-bold">第三階段 (專業)</span>
                <h4 class="text-base font-bold text-white">Power BI 商業智慧</h4>
                <p class="text-xs text-slate-300">建構全航線準點率監控、長約運價比價模型與貨櫃 Demurrage 滯留期成本最佳化儀表板。</p>
              </div>
            </div>

            <div class="p-5 rounded-xl bg-gradient-to-r from-blue-900/40 via-cyan-900/30 to-slate-900 border border-cyan-500/30 text-center">
              <h3 class="text-lg font-bold text-cyan-300">🎉 恭喜完成航運資料前處理核心實戰培訓！</h3>
              <p class="text-xs sm:text-sm text-slate-300 mt-2 max-w-xl mx-auto">
                隨附之完整 Word 教學文件與 30 筆示範檔、300 筆練習檔已在您的工作目錄就緒，歡迎隨時實機演練！
              </p>
            </div>
          </div>
        `;
      }}

      stage.innerHTML = html;

      // Trigger interactive components if needed
      if (slide.content_type === "interactive_dirty_data") {{
        showDirtyDetail(0);
      }} else if (slide.content_type === "interactive_table_toggle") {{
        switchTableMode("raw");
      }} else if (slide.content_type === "interactive_formula_playground") {{
        updateFormulaDemo();
      }} else if (slide.content_type === "interactive_demurrage_calc") {{
        runDemurrageCalc();
      }}
    }}

    // Navigation
    function prevSlide() {{
      if (currentSlide > 0) renderSlide(currentSlide - 1);
    }}
    function nextSlide() {{
      if (currentSlide < slides.length - 1) renderSlide(currentSlide + 1);
    }}
    function goToSlide(idx) {{
      renderSlide(idx);
      document.getElementById("index-modal").classList.add("hidden");
    }}

    // Toggle Speaker Tips
    function toggleSpeakerTips() {{
      const box = document.getElementById("speaker-tips-box");
      const btn = document.getElementById("tips-btn");
      showTips = !showTips;
      if (showTips) {{
        box.classList.remove("hidden");
        btn.classList.add("text-amber-400");
        btn.classList.remove("text-slate-500");
      }} else {{
        box.classList.add("hidden");
        btn.classList.remove("text-amber-400");
        btn.classList.add("text-slate-500");
      }}
    }}

    // Fullscreen
    function toggleFullScreen() {{
      if (!document.fullscreenElement) {{
        document.documentElement.requestFullscreen().catch(() => {{}});
      }} else {{
        if (document.exitFullscreen) {{
          document.exitFullscreen().catch(() => {{}});
        }}
      }}
    }}

    // Lightbox Modal
    function openImageModal(src, caption) {{
      document.getElementById("modal-img").src = src;
      document.getElementById("modal-caption").innerText = caption || "";
      document.getElementById("image-modal").classList.remove("hidden");
      document.getElementById("image-modal").classList.add("flex");
    }}
    function closeImageModal() {{
      document.getElementById("image-modal").classList.add("hidden");
      document.getElementById("image-modal").classList.remove("flex");
    }}

    // Index Modal
    function toggleIndexModal() {{
      const m = document.getElementById("index-modal");
      if (m.classList.contains("hidden")) {{
        buildIndexGrid();
        m.classList.remove("hidden");
        m.classList.add("flex");
      }} else {{
        m.classList.add("hidden");
        m.classList.remove("flex");
      }}
    }}

    function buildIndexGrid() {{
      const grid = document.getElementById("index-grid");
      grid.innerHTML = "";
      slides.forEach((s, idx) => {{
        const isCurrent = idx === currentSlide;
        const card = document.createElement("div");
        card.className = `p-3 rounded-xl border text-left cursor-pointer transition ${{
          isCurrent 
            ? "bg-blue-600/20 border-cyan-400 text-white shadow-lg" 
            : "bg-slate-800/50 hover:bg-slate-800 border-slate-700 text-slate-300"
        }}`;
        card.onclick = () => goToSlide(idx);
        card.innerHTML = `
          <div class="flex items-center justify-between text-[10px] text-cyan-400 font-mono mb-1">
            <span>SLIDE ${{String(idx + 1).padStart(2, '0')}}</span>
            <span>${{s.chapter}}</span>
          </div>
          <h4 class="text-xs font-bold truncate">${{s.title}}</h4>
        `;
        grid.appendChild(card);
      }});
    }}

    // ----------------------------------------------------
    // INTERACTIVE LOGIC: Slide 7 Dirty Data Details
    // ----------------------------------------------------
    const dirtyDataDetails = [
      {{
        name: "首尾不可見空格 (Phantom Spaces)",
        example: "'  EMC-2026-001  ' 或 'OOLU2345678 '",
        damage: "肉眼看似正常，但 VLOOKUP 比對完全失效，系統回傳 #N/A，導致報關提單查無紀錄！",
        solution: "使用 =TRIM(B2) 函數，或使用【資料剖析】批次剔除前後空白贅字。",
        color: "text-amber-400"
      }},
      {{
        name: "代碼大小寫紊亂 (Case Inconsistency)",
        example: "'twkhh' vs 'TWKHH' 或 'emcu1234567'",
        damage: "樞紐分析表會把大小寫視為兩個完全不同的港口分別統計，造成運載量分流腰斬！",
        solution: "使用 =UPPER(TRIM(D2)) 強制轉大寫，保證與國際 UN/LOCODE 標準 100% 吻合。",
        color: "text-cyan-400"
      }},
      {{
        name: "單欄混雜多屬性 (Non-atomic Values)",
        example: "'EVER GIVEN / 0142-052W' (船名與航次擠在同一格)",
        damage: "無法單獨統計特定船舶的全年趟次，亦無法單獨依航向 (E/W) 篩選。",
        solution: "使用【資料剖析】以 '/' 分割，或游標移至隔壁欄按下【Ctrl + E】快速填入！",
        color: "text-blue-400"
      }},
      {{
        name: "數字儲存為文字 (Numbers Stored as Text)",
        example: "毛重打成 '18,500 kg'，運費打成 '$2,450.00'",
        damage: "無法進行 SUM 加總，狀態列只顯示『計數』而無加總，導致報表運費營收算出 0 元！",
        solution: "以 Ctrl+H 取代掉 'kg' 與 '$'，搭配 =VALUE(...) 函數強制轉回純數值型態。",
        color: "text-emerald-400"
      }},
      {{
        name: "多樣化混亂日期 (Inconsistent Dates)",
        example: "'2026/10/05'、'2026-10-22'、'10-10-2026'、'2026.11.08'",
        damage: "文字型日期無法參與相減運算，無法計算航程天數（ETA - ETD）與滯港超期天數。",
        solution: "選取整欄使用【資料剖析】步驟 3 指定【日期 YMD】，Excel 自動統一轉為日期序列值。",
        color: "text-purple-400"
      }},
      {{
        name: "重複託運提單紀錄 (Duplicate Records)",
        example: "同一提單號碼 MSK-2026-991 建檔 2 次 (Row 5 & Row 12)",
        damage: "出口總櫃數與運費被重複計算，報表營收虛胖，產生嚴重審計財務漏洞！",
        solution: "點擊【資料 -> 移除重複項】，單獨勾選『提單號碼』核心主鍵，系統秒速排重。",
        color: "text-rose-400"
      }},
      {{
        name: "關鍵欄位遺漏值 (Missing Values)",
        example: "託運人名稱或付款狀態儲存格完全空白 (Empty/Null)",
        damage: "樞紐分析表出現 '(空白)'，運費請款責任方不明，造成呆帳風險！",
        solution: "使用 =IF(OR(M2=\"\", ISBLANK(M2)), \"[待補託運人]\", M2) 填補業務審計標記。",
        color: "text-indigo-400"
      }}
    ];

    function showDirtyDetail(idx) {{
      const d = dirtyDataDetails[idx];
      const box = document.getElementById("dirty-detail-box");
      
      // Update tab active classes
      const tabs = document.querySelectorAll("#dirty-tabs .dirty-tab-btn");
      tabs.forEach((t, i) => {{
        if (i === idx) {{
          t.className = "dirty-tab-btn active px-3 py-1.5 rounded-lg text-xs font-bold transition bg-cyan-600 text-white";
        }} else {{
          t.className = "dirty-tab-btn px-3 py-1.5 rounded-lg text-xs font-bold transition bg-slate-800 text-slate-300 hover:bg-slate-700";
        }}
      }});

      box.innerHTML = `
        <div class="space-y-2 border-r border-slate-700/60 pr-2">
          <span class="text-xs text-slate-400 font-bold block">典型航運案例範例</span>
          <div class="p-2.5 rounded-lg bg-slate-900 border border-slate-800 font-mono text-xs ${{d.color}} font-bold">
            ${{d.example}}
          </div>
          <span class="text-xs font-bold text-white block mt-2">${{d.name}}</span>
        </div>

        <div class="space-y-2 border-r border-slate-700/60 pr-2">
          <span class="text-xs text-red-400 font-bold block">⚠️ 破壞性後果剖析</span>
          <p class="text-xs text-slate-300 leading-relaxed">${{d.damage}}</p>
        </div>

        <div class="space-y-2">
          <span class="text-xs text-emerald-400 font-bold block">✅ 標準清洗對策</span>
          <p class="text-xs text-slate-300 leading-relaxed">${{d.solution}}</p>
        </div>
      `;
    }}

    // ----------------------------------------------------
    // INTERACTIVE LOGIC: Slide 9 30-Row Table Toggle
    // ----------------------------------------------------
    const rawDemoRows = [
      ["1", "  EMC-2026-001  ", "EVER GIVEN / 0142-052W", "twkhh", "USLAX", "EMCU1234567 ", "40HC", "18500 kg", "$2,450.00", "2026/10/05", "PAID", "宏達國際貿易"],
      ["2", "WHL-2026-088", "WAN HAI 315 / 0089-012N", "TWKEL", "cnsHa", "WHLU9876543", "20GP", " 12,400 ", "USD 1200", "2026-10-06", "Collect", "台灣造紙工業"],
      ["3", "YML-2026-105 ", "YM WELLBEING / 0214-088E", "twkhh", "NLRTM", "YMLU4567890", "40GP", "21300", "EUR 3,100", "10-10-2026", "paid", "長興材料工業"],
      ["4", "OOCL-2026-302", "OOCL POLAND / 0055-031W", "TWKHH", "USLGB", "OOLU2345678", "40HC", "19,800kg", "$2,600", "2026/10/11", "PREPAID", "和大工業"],
      ["5", "MSK-2026-991", "MAERSK MC-KINNEY / 0411-002W", "TWKEL", "DEHAM", "MSKU3456789", "40HC", "22500", "3400 USD", "2026.10.12", "PAID", "日月光半導體"],
      ["12", "MSK-2026-991", "MAERSK MC-KINNEY / 0411-002W", "TWKEL", "DEHAM", "MSKU3456789", "40HC", "22500", "3400 USD", "2026.10.12", "PAID", "日月光半導體 (重複!)"],
      ["15", "EMC-2026-035", "EVER BRAVE / 0067-088N", "TWKEL", "JPYOK", "EMCU4455667", "20GP", "8900 kg", "$800", "2026/10/23", "PAID", "[空白遺漏值]"],
      ["20", "YML-2026-315", "YM WONDER / 0112-064W", "TWKHH", "DEHAM", "YMLU9900112", "40HC", "20700", "3,350 USD", "28-10-2026", "[空白遺漏值]", "正新橡膠工業"]
    ];

    const cleanDemoRows = [
      ["1", "EMC-2026-001", "EVER GIVEN", "0142-052W", "TWKHH", "USLAX", "EMCU1234567", "40HC", "18,500", "2,450.00", "2026-10-05", "2026-10-22", "17", "PREPAID", "宏達國際貿易"],
      ["2", "WHL-2026-088", "WAN HAI 315", "0089-012N", "TWKEL", "CNSHA", "WHLU9876543", "20GP", "12,400", "1,200.00", "2026-10-06", "2026-10-10", "4", "COLLECT", "台灣造紙工業"],
      ["3", "YML-2026-105", "YM WELLBEING", "0214-088E", "TWKHH", "NLRTM", "YMLU4567890", "40GP", "21,300", "3,100.00", "2026-10-10", "2026-11-08", "29", "PREPAID", "長興材料工業"],
      ["4", "OOCL-2026-302", "OOCL POLAND", "0055-031W", "TWKHH", "USLGB", "OOLU2345678", "40HC", "19,800", "2,600.00", "2026-10-11", "2026-10-28", "17", "PREPAID", "和大工業"],
      ["5", "MSK-2026-991", "MAERSK MC", "0411-002W", "TWKEL", "DEHAM", "MSKU3456789", "40HC", "22,500", "3,400.00", "2026-10-12", "2026-11-15", "34", "PREPAID", "日月光半導體"],
      ["14", "EMC-2026-035", "EVER BRAVE", "0067-088N", "TWKEL", "JPYOK", "EMCU4455667", "20GP", "8,900", "800.00", "2026-10-23", "2026-10-27", "4", "PREPAID", "[待補託運人]"],
      ["19", "YML-2026-315", "YM WONDER", "0112-064W", "TWKHH", "DEHAM", "YMLU9900112", "40HC", "20,700", "3,350.00", "2026-10-28", "2026-12-02", "35", "[待確認款項]", "正新橡膠工業"]
    ];

    function switchTableMode(mode) {{
      const bRaw = document.getElementById("btn-raw");
      const bClean = document.getElementById("btn-clean");
      const area = document.getElementById("table-display-area");

      if (mode === "raw") {{
        bRaw.className = "px-3 py-1.5 rounded-lg text-xs font-bold transition bg-red-600 text-white";
        bClean.className = "px-3 py-1.5 rounded-lg text-xs font-bold transition text-slate-400 hover:text-white";

        let html = `
          <table class="w-full text-left text-xs">
            <thead class="bg-red-950/80 text-red-200 border-b border-red-900 sticky top-0">
              <tr>
                <th class="p-2">項次</th>
                <th class="p-2">提單號碼(有空格)</th>
                <th class="p-2">船名航次(擠同格)</th>
                <th class="p-2">裝貨港(小寫)</th>
                <th class="p-2">貨櫃號碼</th>
                <th class="p-2">毛重(帶單位)</th>
                <th class="p-2">海運費(帶幣別)</th>
                <th class="p-2">預計開航日(混雜)</th>
                <th class="p-2">託運人</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-slate-800 text-slate-300 font-mono">
        `;
        rawDemoRows.forEach(r => {{
          const isDup = r[0] === "12";
          const rowBg = isDup ? "bg-red-500/10 text-red-300 font-bold" : "hover:bg-slate-800/40";
          html += `
            <tr class="${{rowBg}}">
              <td class="p-2 font-bold">${{r[0]}}</td>
              <td class="p-2 text-amber-300">${{r[1]}}</td>
              <td class="p-2 text-rose-300">${{r[2]}}</td>
              <td class="p-2 text-purple-300">${{r[3]}}</td>
              <td class="p-2">${{r[5]}}</td>
              <td class="p-2 text-cyan-300">${{r[7]}}</td>
              <td class="p-2 text-emerald-300">${{r[8]}}</td>
              <td class="p-2 text-orange-300">${{r[9]}}</td>
              <td class="p-2 font-sans">${{r[11]}}</td>
            </tr>
          `;
        }});
        html += `</tbody></table>`;
        area.innerHTML = html;
      }} else {{
        bClean.className = "px-3 py-1.5 rounded-lg text-xs font-bold transition bg-emerald-600 text-white";
        bRaw.className = "px-3 py-1.5 rounded-lg text-xs font-bold transition text-slate-400 hover:text-white";

        let html = `
          <table class="w-full text-left text-xs">
            <thead class="bg-emerald-950/80 text-emerald-200 border-b border-emerald-900 sticky top-0">
              <tr>
                <th class="p-2">項次</th>
                <th class="p-2">標準提單號</th>
                <th class="p-2">船名 (已分離)</th>
                <th class="p-2">航次 (已分離)</th>
                <th class="p-2">裝貨港(大寫)</th>
                <th class="p-2">貨櫃號碼</th>
                <th class="p-2">純毛重(KG)</th>
                <th class="p-2">純運費(USD)</th>
                <th class="p-2">標準開航日</th>
                <th class="p-2">航程天數</th>
                <th class="p-2">託運人</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-slate-800 text-slate-300 font-mono">
        `;
        cleanDemoRows.forEach(r => {{
          html += `
            <tr class="hover:bg-slate-800/40">
              <td class="p-2 font-bold">${{r[0]}}</td>
              <td class="p-2 text-emerald-400 font-bold">${{r[1]}}</td>
              <td class="p-2 text-cyan-300">${{r[2]}}</td>
              <td class="p-2 text-cyan-300">${{r[3]}}</td>
              <td class="p-2 text-purple-300 font-bold">${{r[4]}}</td>
              <td class="p-2 text-slate-200">${{r[6]}}</td>
              <td class="p-2 text-right font-bold text-white">${{r[8]}}</td>
              <td class="p-2 text-right font-bold text-emerald-400">${{r[9]}}</td>
              <td class="p-2 text-amber-300">${{r[10]}}</td>
              <td class="p-2 text-center font-bold text-cyan-300">${{r[12]}} 天</td>
              <td class="p-2 font-sans">${{r[14]}}</td>
            </tr>
          `;
        }});
        html += `</tbody></table>`;
        area.innerHTML = html;
      }}
    }}

    // ----------------------------------------------------
    // INTERACTIVE LOGIC: Slide 15 Formula Playground
    // ----------------------------------------------------
    function updateFormulaDemo() {{
      const input = document.getElementById("test-str-input").value;
      const trimmed = input.trim();
      document.getElementById("res-trim").innerText = `"${{trimmed}}"`;

      const slashIdx = trimmed.indexOf("/");
      if (slashIdx !== -1) {{
        const vessel = trimmed.substring(0, slashIdx).trim();
        const voyage = trimmed.substring(slashIdx + 1).trim();
        document.getElementById("res-vessel").innerText = `"${{vessel}}"`;
        document.getElementById("res-voyage").innerText = `"${{voyage}}"`;
      }} else {{
        document.getElementById("res-vessel").innerText = `"${{trimmed}}" (無斜線)`;
        document.getElementById("res-voyage").innerText = `"" (無斜線)`;
      }}

      document.getElementById("res-upper").innerText = `"${{trimmed.toUpperCase()}}"`;
    }}

    // ----------------------------------------------------
    // INTERACTIVE LOGIC: Slide 21 Demurrage Calculator
    // ----------------------------------------------------
    function runDemurrageCalc() {{
      const ataStr = document.getElementById("calc-ata").value;
      const gateStr = document.getElementById("calc-gate").value;
      const freeDays = parseInt(document.getElementById("calc-free").value, 10);

      const ata = new Date(ataStr);
      const gate = new Date(gateStr);
      const diffTime = gate - ata;
      const stayDays = Math.max(0, Math.round(diffTime / (1000 * 60 * 60 * 24)));
      const overdue = Math.max(0, stayDays - freeDays);

      document.getElementById("out-stay").innerText = `${{stayDays}} 天`;
      document.getElementById("out-overdue").innerText = `${{overdue}} 天`;

      const badge = document.getElementById("demurrage-badge");
      if (overdue > 0) {{
        badge.className = "p-4 rounded-xl text-center border font-bold text-base bg-red-500/20 border-red-500/50 text-red-400 shadow-lg";
        badge.innerHTML = `
          <span>🚨 超期警示：滯港罰款 (Demurrage Alert)</span>
          <span class="text-xs text-red-300 font-normal">已超出免租期 ${{overdue}} 天，請即刻向客戶核算滯港費！</span>
        `;
      }} else {{
        badge.className = "p-4 rounded-xl text-center border font-bold text-base bg-emerald-500/20 border-emerald-500/50 text-emerald-400 shadow-lg";
        badge.innerHTML = `
          <span>✅ 正常放行：在免租期內 (Normal / Free)</span>
          <span class="text-xs text-emerald-300 font-normal">剩餘免租額度 ${{freeDays - stayDays}} 天，無須繳納任何滯港費。</span>
        `;
      }}
    }}

    // Keyboard Shortcuts
    document.addEventListener("keydown", (e) => {{
      if (e.target.tagName === "INPUT" || e.target.tagName === "SELECT") return;

      if (e.key === "ArrowRight" || e.key === "PageDown" || e.key === " ") {{
        e.preventDefault();
        nextSlide();
      }} else if (e.key === "ArrowLeft" || e.key === "PageUp") {{
        e.preventDefault();
        prevSlide();
      }} else if (e.key === "Home") {{
        e.preventDefault();
        goToSlide(0);
      }} else if (e.key === "End") {{
        e.preventDefault();
        goToSlide(slides.length - 1);
      }} else if (e.key === "f" || e.key === "F") {{
        e.preventDefault();
        toggleFullScreen();
      }} else if (e.key === "m" || e.key === "M") {{
        e.preventDefault();
        toggleIndexModal();
      }} else if (e.key === "t" || e.key === "T") {{
        e.preventDefault();
        toggleSpeakerTips();
      }} else if (e.key === "Escape") {{
        closeImageModal();
        document.getElementById("index-modal").classList.add("hidden");
      }}
    }});

    // Initialize
    window.addEventListener("DOMContentLoaded", () => {{
      renderSlide(0);
    }});
  </script>
</body>
</html>
"""

# Write to c:\proj\week12-1\index.html and excel_shipping_tutorial_slides.html
with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
with open("excel_shipping_tutorial_slides.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"Generated index.html and excel_shipping_tutorial_slides.html ({len(html_content)} bytes).")
