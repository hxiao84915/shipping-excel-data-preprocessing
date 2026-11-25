# 航運行政人員 MS Excel 基礎資料分析前處理實戰指南

> **專為航運行政、海運承攬（Freight Forwarder）、船務助理與報關人員量身打造的 Excel 資料前處理實戰教學教材**

🌐 **線上公開簡報網站 (GitHub Pages)**：[https://hxiao84915.github.io/shipping-excel-data-preprocessing/](https://hxiao84915.github.io/shipping-excel-data-preprocessing/)

本專案提供一套完整且高規格的電腦操作教學文件、互動簡報、示範資料與練習題庫，深入探討資料分析核心觀念——**「資料前處理」(Data Preprocessing)**，徹底破除航運數據常見的「髒資料」（Dirty Data）陷阱，協助行政人員建立高可用性的數據資產。

---

## 專案核心產出清單 (Deliverables)

| 檔案名稱 | 格式 | 說明 |
| :--- | :---: | :--- |
| **[航運行政人員_MS_Excel基礎資料前處理實戰手冊.docx](./航運行政人員_MS_Excel基礎資料前處理實戰手冊.docx)** | Word (`.docx`) | 完整詳細圖文教學手冊（含封面、章節目錄、圖表目錄、頁碼、10幅高解析截圖與實務對照表） |
| **[航運行政人員_MS_Excel基礎資料前處理實戰指南.pptx](./航運行政人員_MS_Excel基礎資料前處理實戰指南.pptx)** | PowerPoint (`.pptx`) | 16:9 寬螢幕專業簡報（16頁投影片，內建彩色卡片、圖表與全套講師演講備忘稿） |
| **[航運行政_Excel資料前處理_簡報投影片.html](./航運行政_Excel資料前處理_簡報投影片.html)** | 網頁 (`.html`) | 互動式網頁簡報（內建即時公式模擬器、縮圖總覽、全螢幕播放與燈箱放大功能） |
| **[航運行政Excel資料前處理_示範資料檔_30筆.xlsx](./航運行政Excel資料前處理_示範資料檔_30筆.xlsx)** | Excel (`.xlsx`) | 30 筆示範案例（含待清洗原始資料與清洗完成標準主檔對照工作表） |
| **[航運行政Excel資料前處理_綜合練習題庫_300筆.xlsx](./航運行政Excel資料前處理_綜合練習題庫_300筆.xlsx)** | Excel (`.xlsx`) | 300 筆四大情境實戰練習題庫（含全球港口代碼對照表與完整函數公式指引） |

---

## 章節架構導覽

- **第一章 航運行政與 MS Excel 操作介面精要導覽**
  - GIGO 原理與航運業務隱形成本
  - Excel 視窗工作介面全景解析（功能區、名稱方塊、資料編輯列、狀態列）
  - 航運行政日常最高頻快速鍵清單（`Ctrl+E`, `Ctrl+Shift+L`, `F4`, `Ctrl+Alt+V`）
- **第二章 基礎資料前處理的核心常識與思維**
  - 乾淨資料（Tidy Data）三大原則：獨立列、獨立欄、單一原子值
  - 儲存格資料型態（文字、數值、日期序號、布林）與錯誤代碼（`#N/A`, `#VALUE!`, `#REF!`）
  - 航運實務常見「七大髒資料症狀」與衝擊剖析
  - 資料前處理標準五步驟循環 SOP
- **第三章 30筆航運示範案例：資料前處理全流程逐步圖解操作**
  - 30 筆原始航運事務資料診斷與重複提單排查
  - 清洗流程實戰八大步驟指南（`TRIM`, `UPPER`, 移除重複項, 資料剖析, `SUBSTITUTE`, `VALUE`, `DATE`）
  - 清洗前（Dirty）與清洗後（Clean）成效對照（總重量 484,700 KGS、運費 USD 64,550）
- **第四章 前處理必備功能選項與核心函數公式速查寶典**
  - 四大選單工具：資料剖析、移除重複項、快速填入（Flash Fill）、資料驗證下拉防呆
  - 十大核心函數：`TRIM`, `CLEAN`, `UPPER`, `LEFT/MID/RIGHT`, `FIND`, `SUBSTITUTE`, `VALUE`, `TEXT`, `DATE`, `XLOOKUP`, `IFERROR`
- **第五章 300筆航運情境實戰練習題庫與解題全攻略**
  - **模組 A（100筆）**：海運出口訂艙與裝船清單清理（提單去空格、櫃號規格化、毛重純數值萃取）
  - **模組 B（70筆）**：貨櫃滯期延滯費（Demurrage & Detention, D&D）計費天數與逾期分析（`IF(ISBLANK())`, `MAX(0, 天數-免費期)`)
  - **模組 C（70筆）**：國際貨代應收應付帳單幣別、匯率與金額清理（幣別分離、`ROUND` 折合 TWD、5% 營業稅精算）
  - **模組 D（60筆）**：跨國港口代碼與船期表標準化（`XLOOKUP` 比對聯合國 UN/LOCODE、計算 ETD-ETA 航程天數）
- **第六章 航運資料品質控管清單與參考文獻**
  - 航運行政人員資料前處理三階段 SOP 檢核表（清洗前備份、清洗中驗證、清洗後審查）
  - 國際權威參考文獻與引用來源註記
- **附錄**
  - 全球常用 50 大港口 UN/LOCODE 國際代碼速查表
  - 國際標準貨櫃 ISO 6346 尺寸型號代碼標準對照表

---

## 國際標準規範與引用來源

1. **Microsoft Learn Documentation (2024)**: Excel Functions by Category (Text, Lookup, Date/Time).
2. **UN/CEFACT (UNECE)**: UN/LOCODE: Codes for Trade and Transport Locations (Recommendation No. 16).
3. **ISO 8601:2019**: Representation of dates and times (YYYY-MM-DD).
4. **ISO 6346:2022**: Freight containers — Coding, identification and marking.
5. **Hadley Wickham (2014)**: Tidy Data. Journal of Statistical Software, 59(10).
6. **FIATA (2023)**: Multimodal Transport Guidelines and Standard Documentation.

---

## 授權與維護資訊

- **教材編製**：國際航運行政數位轉型教育小組
- **維護帳號**：`hxiao84915@gmail.com`
- **版本**：Version 1.0 (2026年11月)
