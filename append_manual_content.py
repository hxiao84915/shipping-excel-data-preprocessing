import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import pandas as pd

# Import base builder
from build_full_manual import (
    doc, add_p, add_h1, add_h2, add_h3, add_callout, add_figure, add_table_data,
    COLOR_PRIMARY, COLOR_SECONDARY, COLOR_ACCENT, COLOR_DARK, COLOR_MUTED,
    set_cell_background, set_cell_margins
)

print("Appending comprehensive chapters...")

# ===========================================================================
# COVER PAGE (封面)
# ===========================================================================
# Decorative top spacing
p_top = doc.add_paragraph()
p_top.paragraph_format.space_before = Pt(40)
p_top.paragraph_format.space_after = Pt(10)

# Super-title / Category
p_cat = doc.add_paragraph()
p_cat.paragraph_format.space_before = Pt(0)
p_cat.paragraph_format.space_after = Pt(8)
r_cat = p_cat.add_run("MARITIME LOGISTICS & SHIPPING ADMINISTRATION TRAINING SERIES")
r_cat.font.name = "Arial"
r_cat.font.size = Pt(10.5)
r_cat.font.bold = True
r_cat.font.color.rgb = COLOR_SECONDARY

# Main Title
p_title = doc.add_paragraph()
p_title.paragraph_format.space_before = Pt(6)
p_title.paragraph_format.space_after = Pt(12)
r_title = p_title.add_run("航運行政人員 MS Excel 基礎資料分析\n前處理實戰指南")
r_title.font.name = "Microsoft JhengHei"
r_title.font.size = Pt(26)
r_title.font.bold = True
r_title.font.color.rgb = COLOR_PRIMARY

# Subtitle
p_sub = doc.add_paragraph()
p_sub.paragraph_format.space_before = Pt(0)
p_sub.paragraph_format.space_after = Pt(36)
r_sub = p_sub.add_run("從操作介面認識、髒資料特徵剖析、30筆示範案例循序解析\n到300筆四大實務情境題庫全方位解題指南")
r_sub.font.name = "Microsoft JhengHei"
r_sub.font.size = Pt(13)
r_sub.font.color.rgb = COLOR_MUTED

# Decorative horizontal divider bar
p_bar = doc.add_paragraph()
p_bar.paragraph_format.space_before = Pt(0)
p_bar.paragraph_format.space_after = Pt(40)
r_bar = p_bar.add_run("━" * 48)
r_bar.font.name = "Arial"
r_bar.font.size = Pt(12)
r_bar.font.bold = True
r_bar.font.color.rgb = COLOR_ACCENT

# Metadata Table
meta_tbl = doc.add_table(rows=6, cols=2)
meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_data = [
    ("適用對象", "航運公司行政人員、海運承攬運送業 (Freight Forwarder) OP/文件同仁、船務助理、貨櫃場站行政人員及初學 Excel 數據處理者"),
    ("示範工具", "Microsoft Excel 365 / Excel 2019 / Excel 2021 / Excel 2024"),
    ("教材定位", "專注於「基礎資料前處理」(Data Preprocessing) 之電腦實務操作技能"),
    ("製作單位", "國際航運行政數位轉型教育小組 / 業務資訊化推動專案"),
    ("發布日期", "2026 年 11 月正式編製版"),
    ("文件版本", "Version 1.0 (含隨附 30 筆示範活頁簿與 300 筆情境練習題庫)")
]
for idx, (label, val) in enumerate(meta_data):
    c_lbl = meta_tbl.cell(idx, 0)
    c_val = meta_tbl.cell(idx, 1)
    c_lbl.width = Inches(1.5)
    c_val.width = Inches(4.7)
    set_cell_background(c_lbl, "F2F5F9")
    set_cell_background(c_val, "FFFFFF")
    set_cell_margins(c_lbl, top=80, bottom=80, left=100, right=100)
    set_cell_margins(c_val, top=80, bottom=80, left=100, right=100)
    
    p0 = c_lbl.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after = Pt(0)
    r0 = p0.add_run(label)
    r0.font.name = "Microsoft JhengHei"
    r0.font.size = Pt(9.5)
    r0.font.bold = True
    r0.font.color.rgb = COLOR_PRIMARY
    
    p1 = c_val.paragraphs[0]
    p1.paragraph_format.space_before = Pt(0)
    p1.paragraph_format.space_after = Pt(0)
    r1 = p1.add_run(val)
    r1.font.name = "Microsoft JhengHei"
    r1.font.size = Pt(9.5)
    r1.font.color.rgb = COLOR_DARK

# Page break after Cover
doc.add_page_break()

# ===========================================================================
# TABLE OF CONTENTS & LIST OF FIGURES / TABLES (目錄、圖目錄、表目錄)
# ===========================================================================
p_toc_title = doc.add_paragraph()
p_toc_title.paragraph_format.space_before = Pt(10)
p_toc_title.paragraph_format.space_after = Pt(12)
r_tt = p_toc_title.add_run("本手冊章節目錄 (Table of Contents)")
r_tt.font.name = "Microsoft JhengHei"
r_tt.font.size = Pt(18)
r_tt.font.bold = True
r_tt.font.color.rgb = COLOR_PRIMARY

toc_items = [
    ("第一章 航運行政與 MS Excel 操作介面精要導覽", "1"),
    ("    1.1 航運行政人員為什麼必須精通「資料前處理」？", "1"),
    ("    1.2 Excel 365 視窗工作介面全景解析與重要元素說明", "2"),
    ("    1.3 航運行政日常最高頻「功能區索引標籤」功能導覽", "3"),
    ("    1.4 航運文書處理必備快速鍵速查清單", "4"),
    ("第二章 基礎資料前處理的核心常識與思維", "5"),
    ("    2.1 什麼是「乾淨資料」(Tidy Data)？航運數據三大基本原則", "5"),
    ("    2.2 儲存格資料型態原理與實務陷阱", "6"),
    ("    2.3 航運實務常見「七大髒資料症狀」與衝擊剖析", "7"),
    ("    2.4 資料前處理標準五步驟循環 SOP", "9"),
    ("第三章 30筆航運示範案例：資料前處理全流程逐步圖解操作", "11"),
    ("    3.1 30筆原始航運事務資料背景與欄位結構說明", "11"),
    ("    3.2 示範案例清洗成效總覽對照", "13"),
    ("    3.3 步驟一：工作表備份與欄位格式初步檢查", "14"),
    ("    3.4 步驟二：提單號碼多餘空白消除與大小寫轉換 (TRIM, UPPER)", "14"),
    ("    3.5 步驟三：重疊與重複提單記錄之比對與移除 (移除重複項)", "15"),
    ("    3.6 步驟四：船名與航次複合欄位之原子化拆分 (資料剖析 / 閃電填入)", "16"),
    ("    3.7 步驟五：起訖港口代碼與貨櫃規格標準化處理 (UN/LOCODE, ISO規格)", "18"),
    ("    3.8 步驟六：包裝件數、毛重與材積之數值與單位分離 (SUBSTITUTE, VALUE)", "19"),
    ("    3.9 步驟七：非標準日期格式統一為西元標準日 (DATE, 資料剖析)", "21"),
    ("    3.10 步驟八：運費條款標準化與金額純淨轉型", "22"),
    ("    3.11 30筆示範案例清洗完成標準成果表", "23"),
    ("第四章 前處理必備功能選項與核心函數公式速查寶典", "25"),
    ("    4.1 功能區選單操作手冊 (資料剖析、移除重複、閃電填入、資料驗證)", "25"),
    ("    4.2 航運資料前處理「十大核心函數」語法與精選範例", "28"),
    ("    4.3 航運業務複合巢狀公式進階解題示範", "33"),
    ("第五章 300筆航運情境實戰練習題庫與解題全攻略", "35"),
    ("    5.1 300筆實戰題庫四大業務情境架構導覽", "35"),
    ("    5.2 模組 A：海運出口訂艙與裝船清單清理 (100筆)", "36"),
    ("    5.3 模組 B：貨櫃滯期延滯費 (D&D) 計費天數與逾期分析前處理 (70筆)", "40"),
    ("    5.4 模組 C：國際貨代應收應付帳單幣別、匯率與金額清理 (70筆)", "44"),
    ("    5.5 模組 D：跨國港口代碼與船期表標準化 (60筆)", "48"),
    ("    5.6 隨附活頁簿結構說明與練習驗收方式", "52"),
    ("第六章 航運資料品質控管清單與參考文獻", "53"),
    ("    6.1 航運行政人員資料前處理三階段 SOP 檢核表", "53"),
    ("    6.2 參考文獻與國際海運標準規範註記", "55"),
    ("附錄：實用代碼速查表與資源指南", "56"),
    ("    附錄一：全球常用 50 大港口 UN/LOCODE 國際代碼速查表", "56"),
    ("    附錄二：國際 ISO 6346 貨櫃尺寸型號代碼標準對照表", "58"),
    ("    附錄三：隨附 Excel 練習檔工作表索引說明", "59")
]

for title, page_str in toc_items:
    p_item = doc.add_paragraph()
    p_item.paragraph_format.space_before = Pt(1)
    p_item.paragraph_format.space_after = Pt(2)
    p_item.paragraph_format.line_spacing = 1.15
    
    # Leader dots calculation
    r1 = p_item.add_run(title)
    r1.font.name = "Microsoft JhengHei"
    r1.font.size = Pt(9.5)
    if not title.startswith("    "):
        r1.font.bold = True
        r1.font.color.rgb = COLOR_PRIMARY
    else:
        r1.font.color.rgb = COLOR_DARK
        
    dots_count = max(5, 75 - len(title) * 2)
    r_dots = p_item.add_run(" " + "." * dots_count + " ")
    r_dots.font.name = "Arial"
    r_dots.font.size = Pt(8.5)
    r_dots.font.color.rgb = RGBColor(180, 180, 180)
    
    r_p = p_item.add_run(page_str)
    r_p.font.name = "Arial"
    r_p.font.size = Pt(9.5)
    r_p.font.bold = True
    r_p.font.color.rgb = COLOR_PRIMARY

# Add List of Figures
p_fig_title = doc.add_paragraph()
p_fig_title.paragraph_format.space_before = Pt(14)
p_fig_title.paragraph_format.space_after = Pt(6)
r_ft = p_fig_title.add_run("圖目錄 (List of Figures)")
r_ft.font.name = "Microsoft JhengHei"
r_ft.font.size = Pt(13)
r_ft.font.bold = True
r_ft.font.color.rgb = COLOR_PRIMARY

fig_list = [
    ("圖 1-1  Excel 365 視窗工作介面全景導覽與前處理核心工具標註", "2"),
    ("圖 2-1  航運資料前處理核心五步循環 SOP 流程圖", "9"),
    ("圖 2-2  航運實務常見『七大髒資料症狀』與標準清洗對策架構", "8"),
    ("圖 3-1  30筆航運示範案例清洗前 (Dirty) 與 清洗後 (Clean) 成效對照圖", "13"),
    ("圖 4-1  Excel 資料剖析精靈步驟 2/3 分隔符號設定對話方塊", "17"),
    ("圖 4-2  Excel 移除重複項 (Remove Duplicates) 功能選項設定視窗", "16"),
    ("圖 4-3  Excel「快速填入」(Flash Fill / Ctrl+E) 智慧拆解示範畫面", "18"),
    ("圖 4-4  Excel 資料驗證 (Data Validation) 貨櫃規格下拉防呆設定", "27"),
    ("圖 4-5  航運資料前處理『十大核心函數』體系架構導圖", "28"),
    ("圖 5-1  300筆航運情境練習題庫四大業務模組架構圖", "35")
]
for f_title, f_p_str in fig_list:
    p_f = doc.add_paragraph()
    p_f.paragraph_format.space_before = Pt(1)
    p_f.paragraph_format.space_after = Pt(2)
    rf = p_f.add_run(f_title)
    rf.font.name = "Microsoft JhengHei"
    rf.font.size = Pt(9)
    rf.font.color.rgb = COLOR_DARK
    
    dots_c = max(5, 72 - len(f_title) * 2)
    rf_dots = p_f.add_run(" " + "." * dots_c + " ")
    rf_dots.font.name = "Arial"
    rf_dots.font.size = Pt(8.5)
    rf_dots.font.color.rgb = RGBColor(180, 180, 180)
    
    rf_p = p_f.add_run(f_p_str)
    rf_p.font.name = "Arial"
    rf_p.font.size = Pt(9)
    rf_p.font.color.rgb = COLOR_MUTED

# Add List of Tables
p_tbl_title = doc.add_paragraph()
p_tbl_title.paragraph_format.space_before = Pt(12)
p_tbl_title.paragraph_format.space_after = Pt(6)
r_tb = p_tbl_title.add_run("表目錄 (List of Tables)")
r_tb.font.name = "Microsoft JhengHei"
r_tb.font.size = Pt(13)
r_tb.font.bold = True
r_tb.font.color.rgb = COLOR_PRIMARY

tbl_list = [
    ("表 1-1  航運行政日常高頻快速鍵速查表", "4"),
    ("表 2-1  Excel 五大基礎資料型態與運算特性對照表", "6"),
    ("表 2-2  Excel 常見錯誤訊息代碼原因與航運對策", "7"),
    ("表 3-1  30筆航運事務示範案例原始資料表 (待清洗髒資料)", "11"),
    ("表 3-2  30筆航運示範案例資料清洗完成標準主檔 (29筆唯一記錄)", "23"),
    ("表 4-1  前處理十大核心函數語法、用途與航運實務範例總覽", "32"),
    ("表 5-1  300筆實務題庫模組 A：出口訂艙裝船清單原始結構 (前10筆示範)", "37"),
    ("表 5-2  300筆實務題庫模組 B：貨櫃滯期延滯費 (D&D) 計算結構", "41"),
    ("表 5-3  300筆實務題庫模組 C：貨代帳單幣別匯率金額清理結構", "45"),
    ("表 5-4  300筆實務題庫模組 D：跨國港口與船期表標準化結構", "49"),
    ("表 6-1  航運行政人員資料前處理三階段 SOP 檢核查核表", "53"),
    ("表 附錄-1  全球 50 大主要貨櫃港口 UN/LOCODE 代碼速查對照表", "56"),
    ("表 附錄-2  國際標準貨櫃 ISO 6346 尺寸型號代碼標準對照表", "58")
]
for t_title, t_p_str in tbl_list:
    p_t = doc.add_paragraph()
    p_t.paragraph_format.space_before = Pt(1)
    p_t.paragraph_format.space_after = Pt(2)
    rt = p_t.add_run(t_title)
    rt.font.name = "Microsoft JhengHei"
    rt.font.size = Pt(9)
    rt.font.color.rgb = COLOR_DARK
    
    dots_c = max(5, 72 - len(t_title) * 2)
    rt_dots = p_t.add_run(" " + "." * dots_c + " ")
    rt_dots.font.name = "Arial"
    rt_dots.font.size = Pt(8.5)
    rt_dots.font.color.rgb = RGBColor(180, 180, 180)
    
    rt_p = p_t.add_run(t_p_str)
    rt_p.font.name = "Arial"
    rt_p.font.size = Pt(9)
    rt_p.font.color.rgb = COLOR_MUTED

print("Cover and TOC generated.")
