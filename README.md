# 航運行政人員 MS Excel 基礎資料前處理操作實戰指引
> **專為海運貨代、船務行政、港埠營運人員打造的電腦技能速成教材與大數據練習套件**

---

## 📌 專案簡介 (Overview)

本專案旨在提供一份完整、系統化且貼近航運業界真實情境的 **MS Excel 基礎資料分析前處理（Data Preprocessing）教學資源庫**。
針對初入航運行政、船務助理、貨運代理（Freight Forwarder OP）、報關助理及港埠物流人員，解決日常提單（B/L）、訂艙單（Booking）、船期表、運費清單及櫃態追蹤表中最棘手的「骯髒資料（Dirty Data）」問題。

遵循國際資料管理協會（DAMA-DMBOK2）六大資料品質維度、關聯式資料庫第一正規化（1NF / Tidy Data）規範，以及 ISO 6346 貨櫃編碼、UN/LOCODE 聯合國港口代碼國際標準。

---

## 📂 專案檔案清單 (Repository Structure)

```text
├── 航運行政人員_MS_Excel基礎資料前處理操作實戰指引.pptx   # 23 頁 16:9 高清 PowerPoint 簡報 (含演講者備忘稿)
├── 航運行政人員_MS_Excel基礎資料前處理操作實戰指引.docx   # 完整版 MS Word 教學手冊 (含封面、目錄、圖表目錄與頁碼)
├── excel_shipping_tutorial_slides.html                    # 互動式 PPT 投影片網頁 (同 index.html)
├── index.html                                             # 網頁版首頁 (支援 GitHub Pages 一鍵展示)
├── Shipping_Demo_30.xlsx                                  # 30 筆真實航運示範資料檔 (Raw / Cleaned / Rules)
├── Shipping_Practice_300.xlsx                             # 300 筆大數據航運實務練習檔 (Raw / Master / Tasks / Solutions)
├── images/                                                # 12 幅教學圖表、Excel 介面截圖與對話方塊示意圖
│   ├── fig1_excel_interface.png                           # 圖 1-1：Excel 核心介面八大功能區導覽
│   ├── fig2_data_hygiene_good_vs_bad.png                  # 圖 2-1：錯誤排版 vs 分析型標準二維表
│   ├── fig3_flash_fill_ctrl_e.png                         # 圖 3-1：快速填入 Ctrl+E 智慧萃取
│   ├── fig4_text_to_columns_wizard.png                    # 圖 3-2：資料剖析三步驟精靈
│   ├── fig5_remove_duplicates_dialog.png                  # 圖 3-3：移除重複項對話框
│   ├── fig6_find_and_replace_dialog.png                   # 圖 3-4：尋找與取代 Ctrl+H
│   ├── fig7_data_validation_dialog.png                    # 圖 3-5：資料驗證下拉選單
│   ├── fig8_filter_and_sort_ui.png                        # 圖 3-6：自動篩選選單
│   ├── fig9_text_formulas_workflow.png                    # 圖 3-7：文字清洗五大王牌函數鏈
│   ├── fig10_date_functions_and_transit.png               # 圖 3-8：日期轉換與航程天數運算
│   ├── fig11_xlookup_port_master.png                      # 圖 4-2：XLOOKUP 跨表港口關聯
│   └── fig12_practice_300_workflow.png                    # 圖 4-1：300 筆大數據清洗管線
└── build scripts/                                         # 自動化編譯與生成 Python 腳本
    ├── generate_datasets.py                               # 產生 30 筆與 300 筆 Excel 檔
    ├── generate_screenshots.py                            # 產生 12 幅教學圖表
    ├── build_docx.py                                      # 編譯完整 Word 教學文件
    ├── build_interactive_presentation.py                  # 編譯互動式 PPT 網頁
    └── build_pptx.py                                      # 編譯 23 頁 PowerPoint 簡報檔
```

---

## 🎯 核心教學模組 (Modules)

### 第一章：Excel 操作介面與新手導航
- **核心架構**：標題列、功能區索引標籤（常用、資料、公式）、名稱方塊、資料編輯列、儲存格網格與狀態列。
- **航運必備十大高頻快捷鍵**：`Ctrl+Shift+L` 自動篩選、`Ctrl+E` 快速填入、`Ctrl+H` 取代、`Ctrl+1` 格式設定、`Alt+=` 自動加總等。

### 第二章：資料前處理應具備的基本常識
- **垃圾進、垃圾出 (GIGO 原則)**：前處理佔據分析 80% 時間的二八定律。
- **第一正規化 (1NF / Tidy Data)**：單一列標頭、一列一紀錄、一欄一原子屬性、**絕對嚴禁合併儲存格**。
- **航運七大骯髒資料 (Dirty Data) 剖析**：
  1. 首尾不可見空格（`"  EMC-2026-001  "`）
  2. 代碼大小寫紊亂（`twkhh` vs `TWKHH`）
  3. 單欄混雜多屬性（`EVER GIVEN / 0142-052W`）
  4. 數字存為文字（`18,500 kg`、`$2,450.00`）
  5. 多樣化混亂日期（`2026/10/05`、`2026-10-22`、`10-10-2026`、`2026.11.08`）
  6. 重複建單紀錄（提單 `MSK-2026-991` 重複建單）
  7. 關鍵欄位遺漏值（託運人或付款狀態空白）
- **四大本質資料型態**：文字（靠左）、數值（靠右）、日期序列值（本質為天數）、布林值，以及綠色錯誤三角形之排錯。

### 第三章：30 筆真實航運資料實戰示範
- **UI 功能選項**：
  - 【資料剖析】：三步驟將「船名航次」無損拆分。
  - 【快速填入 (Ctrl+E)】：建立範本，一秒智慧萃取。
  - 【移除重複項】：以提單號碼（B/L No）為唯一業務主鍵排重。
  - 【尋找與取代 (Ctrl+H)】：批次拔除毛重 `' kg'` 與幣別 `'$'`。
  - 【資料驗證】：建立 `PREPAID, COLLECT` 下拉清單源頭防錯。
  - 【自動篩選 (Ctrl+Shift+L)】：清單聚合唯一值，快速嗅探異常。
- **核心函數公式鏈**：
  - 文字清洗鏈：`=TRIM()`、`=UPPER()`、`=LEFT(..., FIND()-1)`、`=MID(..., FIND()+1, 50)`。
  - 數值淨化：`=VALUE(SUBSTITUTE(SUBSTITUTE(I2, "$", ""), "USD", ""))`。
  - 日期與航期運算：`=DATEVALUE(...)` 與 `=ETA - ETD` 計算航程天數。
  - 遺漏值防錯：`=IF(OR(M2="", ISBLANK(M2)), "[待補託運人]", TRIM(M2))`。

### 第四章：300 筆大數據航運實務案例情境全解
涵蓋遠洋（美西/美東/西北歐）、近洋與東南亞主要航線，跨越長榮、陽明、萬海、馬士基、達飛、ONE 等航商：
- **核心碼表關聯**：
  - 港口主檔（Port Master）：UN/LOCODE（TWKHH、DEHAM、NLRTM、USLAX 等）。
  - ISO 6346 櫃型主檔：20GP (1.0 TEU)、40GP (2.0 TEU)、40HC (2.0 TEU)、45HC (2.25 TEU) 等。
- **八大實務任務挑戰 (Task 1 ~ Task 8)**：
  - 任務 1：提單與訂艙編號標準化 (`=TRIM(UPPER(C2))`)
  - 任務 2：裝卸港 UN/LOCODE 跨表關聯 (`=XLOOKUP(...)` / `=VLOOKUP(...)`)
  - 任務 3：貨櫃號碼與 ISO 櫃型 TEU 換算
  - 任務 4：毛重與包裝單位純數值化
  - 任務 5：海運費幣別分離與總費用試算
  - 任務 6：多樣日期清洗與航期計算
  - 任務 7：港口免租期與超期滯港費 (Demurrage) 預警：
    $$\text{超期天數} = \max\Big(0,\; (\text{Gate-Out} - \text{ATA}) - \text{Free Days}\Big)$$
    公式：`=MAX(0, (GateOut - ATA) - FreeDays)` 搭配 `=IF(..., "滯港罰款", "正常")`
  - 任務 8：全域重複提單排查與缺漏資料稽核

### 第五章：參考文獻與王牌函數速查手冊
- 引用 ISO 6346、UN/CEFACT UN/LOCODE、DAMA-DMBOK2、Hadley Wickham Tidy Data 與 Microsoft 官方文獻。
- 提供十大王牌函數標準語法速查表與通往 Power Query、Power BI 的三階段學習藍圖。

---

## 🚀 互動式網頁與簡報特色 (Interactive Presentation)

開啟 `index.html` 或 `excel_shipping_tutorial_slides.html`，即可體驗專為演講與自學打造的現代化 PPT 介面：
- **鍵盤與簡報筆操控**：`←` / `→` 切換、`空白鍵` 下一張、`F` 全螢幕、`M` 目錄跳轉、`T` 講者備忘。
- **文字清洗試驗場 (Formula Playground)**：在線自由輸入任意混亂字串，即時查看公式拆解結果。
- **滯港費 (Demurrage) 即時運算機**：動態選擇到港日、提櫃日與免租期，即時觸發紅綠燈預警。
- **圖表全螢幕燈箱**：點擊任一圖表即可全螢幕無損放大檢視。

---

## 📜 授權與引用說明 (License & Citation)
- 編撰單位：航運數位營運與商業分析培訓小組
- 版本：Version 1.0 (2026年11月)
- 適用環境：Microsoft Excel 365 / Excel 2021+
