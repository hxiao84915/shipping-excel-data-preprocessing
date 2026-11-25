import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

print("Starting to build PowerPoint presentation...")

prs = Presentation()
# Set widescreen 16:9 (13.333 x 7.5 inches)
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

blank_slide_layout = prs.slide_layouts[6]

# Theme Colors
BG_COLOR = RGBColor(11, 21, 40)        # Deep Navy #0B1528
CARD_BG = RGBColor(18, 35, 68)         # #122344
CARD_BORDER = RGBColor(40, 60, 100)    # #283C64
TEXT_WHITE = RGBColor(255, 255, 255)
TEXT_MUTED = RGBColor(148, 163, 184)   # Slate-400
TEXT_CYAN = RGBColor(0, 210, 255)      # Cyan #00D2FF
TEXT_EMERALD = RGBColor(16, 185, 129)  # Emerald #10B981
TEXT_AMBER = RGBColor(245, 158, 11)    # Amber #F59E0B
TEXT_ROSE = RGBColor(244, 63, 94)      # Rose #F43F5E

fig_dir = "c:\\proj\\w12\\figures"

def set_slide_background(slide):
    # Add dark background rectangle covering whole slide
    bg_box = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height
    )
    bg_box.fill.solid()
    bg_box.fill.fore_color.rgb = BG_COLOR
    bg_box.line.fill.background() # no line

def add_header(slide, category, title, slide_num):
    # Top bar
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.9))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    # Category / breadcrumb
    p0 = tf.paragraphs[0]
    r0 = p0.add_run()
    r0.text = f"● {category.upper()}   |   航運行政 MS EXCEL 基礎資料分析前處理實戰指南"
    r0.font.name = "Microsoft JhengHei"
    r0.font.size = Pt(9.5)
    r0.font.bold = True
    r0.font.color.rgb = TEXT_CYAN
    
    # Title
    p1 = tf.add_paragraph()
    p1.space_before = Pt(3)
    r1 = p1.add_run()
    r1.text = title
    r1.font.name = "Microsoft JhengHei"
    r1.font.size = Pt(20)
    r1.font.bold = True
    r1.font.color.rgb = TEXT_WHITE
    
    # Slide Number pill (top right)
    pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.6), Inches(0.45), Inches(0.933), Inches(0.35))
    pill.fill.solid()
    pill.fill.fore_color.rgb = RGBColor(15, 29, 56)
    pill.line.color.rgb = CARD_BORDER
    p_pill = pill.text_frame.paragraphs[0]
    p_pill.alignment = PP_ALIGN.CENTER
    r_pill = p_pill.add_run()
    r_pill.text = f"{slide_num:02d} / 16"
    r_pill.font.name = "Arial"
    r_pill.font.size = Pt(10)
    r_pill.font.bold = True
    r_pill.font.color.rgb = TEXT_CYAN

def add_footer(slide):
    # Subtle bottom border line
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(7.05), Inches(11.733), Inches(0.015))
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(30, 45, 75)
    line.line.fill.background()
    
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(7.1), Inches(11.733), Inches(0.3))
    tf = tb.text_frame
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "國際航運行政數位轉型教育小組 編製  •  隨附 30 筆示範活頁簿與 300 筆實戰情境題庫  •  符合 UN/LOCODE 與 ISO 6346 標準"
    r.font.name = "Microsoft JhengHei"
    r.font.size = Pt(8.5)
    r.font.color.rgb = TEXT_MUTED

def set_speaker_notes(slide, notes_text):
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = notes_text

# ===========================================================================
# SLIDE 1: 封面投影片
# ===========================================================================
slide1 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(slide1)

# Large Center Title Box
tb_cov = slide1.shapes.add_textbox(Inches(1.2), Inches(1.3), Inches(10.933), Inches(3.8))
tf_cov = tb_cov.text_frame
tf_cov.word_wrap = True

# Super title tag
p_t0 = tf_cov.paragraphs[0]
p_t0.alignment = PP_ALIGN.CENTER
r_t0 = p_t0.add_run()
r_t0.text = "MARITIME LOGISTICS & SHIPPING DATA PREPROCESSING"
r_t0.font.name = "Arial"
r_t0.font.size = Pt(12)
r_t0.font.bold = True
r_t0.font.color.rgb = TEXT_CYAN

# Main title
p_t1 = tf_cov.add_paragraph()
p_t1.space_before = Pt(12)
p_t1.alignment = PP_ALIGN.CENTER
r_t1 = p_t1.add_run()
r_t1.text = "航運行政人員 MS Excel 基礎資料分析\n前處理實戰指南"
r_t1.font.name = "Microsoft JhengHei"
r_t1.font.size = Pt(36)
r_t1.font.bold = True
r_t1.font.color.rgb = TEXT_WHITE

# Subtitle
p_t2 = tf_cov.add_paragraph()
p_t2.space_before = Pt(14)
p_t2.alignment = PP_ALIGN.CENTER
r_t2 = p_t2.add_run()
r_t2.text = "從操作介面認識、常見髒資料特徵剖析、30筆示範案例全流程逐步圖解\n到300筆四大航運實務情境題庫全攻略（出口清單 / 貨櫃滯期費 / 貨代帳單 / 港口船期）"
r_t2.font.name = "Microsoft JhengHei"
r_t2.font.size = Pt(13)
r_t2.font.color.rgb = TEXT_MUTED

# 4 Badges bottom
b_w = Inches(2.6)
b_h = Inches(1.1)
b_y = Inches(5.3)
badges = [
    ("適用對象", "航運/承攬/報關行政初學", TEXT_WHITE),
    ("示範工具", "MS Excel 365 / 2021 / 2024", TEXT_WHITE),
    ("實例演練", "30筆示範 + 300筆實戰", TEXT_CYAN),
    ("國際規範", "UN/LOCODE & ISO 6346", TEXT_EMERALD)
]
for i, (b_title, b_val, col) in enumerate(badges):
    b_x = Inches(1.2 + i * 2.8)
    card = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, b_x, b_y, b_w, b_h)
    card.fill.solid()
    card.fill.fore_color.rgb = CARD_BG
    card.line.color.rgb = CARD_BORDER
    c_tf = card.text_frame
    c_tf.word_wrap = True
    c_tf.margin_top = Inches(0.18)
    
    cp0 = c_tf.paragraphs[0]
    cp0.alignment = PP_ALIGN.CENTER
    cr0 = cp0.add_run()
    cr0.text = b_title
    cr0.font.name = "Microsoft JhengHei"
    cr0.font.size = Pt(9.5)
    cr0.font.color.rgb = TEXT_MUTED
    
    cp1 = c_tf.add_paragraph()
    cp1.alignment = PP_ALIGN.CENTER
    cr1 = cp1.add_run()
    cr1.text = b_val
    cr1.font.name = "Microsoft JhengHei"
    cr1.font.size = Pt(11)
    cr1.font.bold = True
    cr1.font.color.rgb = col

set_speaker_notes(slide1, "歡迎進入航運行政人員 Excel 前處理實戰課程。本手冊專為初學者量身打造，聚焦於最耗時且最容易出錯的『資料前處理』，從底層邏輯打好數據分析基本功。")

# ===========================================================================
# SLIDE 2: 為什麼航運行政人員必須精通「資料前處理」？
# ===========================================================================
slide2 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(slide2)
add_header(slide2, "第一章 介面與導覽", "1.1 為什麼航運行政人員必須精通「資料前處理」？", 2)
add_footer(slide2)

# Subtitle banner
tb_s2 = slide2.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(11.733), Inches(0.4))
tb_s2.text_frame.paragraphs[0].text = "數據分析的黃金鐵律：『垃圾進，垃圾出』(Garbage In, Garbage Out, GIGO)"
tb_s2.text_frame.paragraphs[0].font.name = "Microsoft JhengHei"
tb_s2.text_frame.paragraphs[0].font.size = Pt(12)
tb_s2.text_frame.paragraphs[0].font.color.rgb = TEXT_MUTED

# 3 Pain Point Cards
c_w = Inches(3.7)
c_h = Inches(3.6)
c_y = Inches(1.9)
cards_s2 = [
    ("1. 提單號碼比對失敗 (#N/A)", 
     "提單前後夾帶半形空格（如 \" EMCU20491823 \"），VLOOKUP/XLOOKUP 直接判定為不同字串，導致漏單、無法核對貨櫃動態。\n\n衝擊後果：\n• 海運發單嚴重延誤\n• 系統拋出大量 #N/A 錯誤\n• 人工逐筆比對耗費數小時", 
     TEXT_ROSE),
    ("2. 貨物重量加總為 0", 
     "貨物毛重輸入為 \"18,450.50 KGS\"，儲存格被 Excel 強制判定為「純文字」而非「數值」，SUM 加總函數直接忽略顯示為 0。\n\n衝擊後果：\n• 船舶配載重心計算失準\n• 應收運費短計漏收\n• 樞紐分析表統計失靈", 
     TEXT_AMBER),
    ("3. 滯期天數與航期算錯", 
     "開航日登記為 \"2026.11.05\" 等文字格式，導致兩日期無法進行減法天數運算，貨櫃延滯費（D&D）免收期判定錯誤。\n\n衝擊後果：\n• 產生巨額場站滯期費爭議\n• 跨月帳單核銷對不上\n• 客戶關係受損與索賠", 
     TEXT_CYAN)
]

for idx, (ctitle, cdesc, ccol) in enumerate(cards_s2):
    cx = Inches(0.8 + idx * 4.0)
    box = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, c_y, c_w, c_h)
    box.fill.solid()
    box.fill.fore_color.rgb = CARD_BG
    box.line.color.rgb = ccol
    box.line.width = Pt(1.5)
    
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_top = Inches(0.2)
    tf.margin_left = Inches(0.25)
    tf.margin_right = Inches(0.25)
    
    p0 = tf.paragraphs[0]
    r0 = p0.add_run()
    r0.text = ctitle
    r0.font.name = "Microsoft JhengHei"
    r0.font.size = Pt(13)
    r0.font.bold = True
    r0.font.color.rgb = ccol
    
    p1 = tf.add_paragraph()
    p1.space_before = Pt(10)
    r1 = p1.add_run()
    r1.text = cdesc
    r1.font.name = "Microsoft JhengHei"
    r1.font.size = Pt(10.5)
    r1.font.color.rgb = TEXT_WHITE

# Bottom Tip Box
tip_box = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.75), Inches(11.733), Inches(1.05))
tip_box.fill.solid()
tip_box.fill.fore_color.rgb = RGBColor(12, 35, 60)
tip_box.line.color.rgb = TEXT_CYAN
tip_tf = tip_box.text_frame
tip_tf.word_wrap = True
tip_tf.margin_top = Inches(0.18)
tip_tf.margin_left = Inches(0.3)
tip_p = tip_tf.paragraphs[0]
tip_r0 = tip_p.add_run()
tip_r0.text = "★ 關鍵認知：資料前處理佔據數據分析 80% 的工作時間！\n"
tip_r0.font.name = "Microsoft JhengHei"
tip_r0.font.size = Pt(11)
tip_r0.font.bold = True
tip_r0.font.color.rgb = TEXT_CYAN
tip_r1 = tip_p.add_run()
tip_r1.text = "只有先清洗出高品質的乾淨資料（Tidy Data），後續的統計運算、樞紐分析表（Pivot Table）與視覺化儀表板才能發揮真正威力。"
tip_r1.font.name = "Microsoft JhengHei"
tip_r1.font.size = Pt(10)
tip_r1.font.color.rgb = TEXT_WHITE

set_speaker_notes(slide2, "向學員剖析航運業務中最常見的三大痛點：提單前後多空格導致VLOOKUP爆掉、重量帶單位導致加總為0、日期點號導致無法算滯期天數。這些都是未做資料前處理的代價。")

# ===========================================================================
# SLIDE 3: Excel 365 視窗工作介面全景解析
# ===========================================================================
slide3 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(slide3)
add_header(slide3, "第一章 介面與導覽", "1.2 Excel 365 視窗工作介面全景導覽與核心元素", 3)
add_footer(slide3)

# Left Image Box
img_path_1 = os.path.join(fig_dir, "fig1_excel_interface.png")
if os.path.exists(img_path_1):
    slide3.shapes.add_picture(img_path_1, Inches(0.8), Inches(1.5), Inches(7.2), Inches(4.5))

# Right 4 Key Elements
e_w = Inches(4.2)
e_h = Inches(1.0)
elements = [
    ("1. 功能區 Ribbon 與「資料」標籤", "集合前處理必備的【資料工具群】（資料剖析、移除重複項、閃電填入、資料驗證）。", TEXT_CYAN),
    ("2. 名稱方塊 (Name Box)", "顯示游標座標（如 B2）。在數千列大表中可直接鍵入目標座標秒速跳轉定位。", TEXT_WHITE),
    ("3. 資料編輯列 (Formula Bar)", "忠實顯示儲存格底層真實內容（公式或未格式化純值），手動編輯與除錯專用。", TEXT_EMERALD),
    ("4. 狀態列 (Status Bar)", "選取範圍時即時顯示計數（Count）、加總（Sum）與平均值，快速核對總重量。", TEXT_AMBER)
]
for idx, (etitle, edesc, ecol) in enumerate(elements):
    ey = Inches(1.5 + idx * 1.35)
    box = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.3), ey, e_w, Inches(1.2))
    box.fill.solid()
    box.fill.fore_color.rgb = CARD_BG
    box.line.color.rgb = ecol
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_top = Inches(0.12)
    tf.margin_left = Inches(0.2)
    
    p0 = tf.paragraphs[0]
    r0 = p0.add_run()
    r0.text = etitle
    r0.font.name = "Microsoft JhengHei"
    r0.font.size = Pt(11)
    r0.font.bold = True
    r0.font.color.rgb = ecol
    
    p1 = tf.add_paragraph()
    p1.space_before = Pt(3)
    r1 = p1.add_run()
    r1.text = edesc
    r1.font.name = "Microsoft JhengHei"
    r1.font.size = Pt(9.5)
    r1.font.color.rgb = TEXT_WHITE

# Bottom Tip
b_box3 = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.15), Inches(11.733), Inches(0.75))
b_box3.fill.solid()
b_box3.fill.fore_color.rgb = CARD_BG
b_box3.line.color.rgb = CARD_BORDER
b_tf3 = b_box3.text_frame
b_tf3.margin_top = Inches(0.12)
b_tf3.margin_left = Inches(0.25)
b_p3 = b_tf3.paragraphs[0]
b_p3.text = "💡 操作提示：功能區「資料」標籤是前處理最重要的核心基地；底端狀態列右鍵可自訂勾選最大值、最小值，是免寫公式快速驗算的利器。"
b_p3.font.name = "Microsoft JhengHei"
b_p3.font.size = Pt(10)
b_p3.font.color.rgb = TEXT_CYAN

set_speaker_notes(slide3, "引導學員觀察圖1-1各區域。特別強調資料編輯列會顯示儲存格底層真實公式，例如看到的是乾淨提單，但編輯列會顯示 =UPPER(TRIM(A2))。")

# ===========================================================================
# SLIDE 4: 航運文書處理必備快速鍵速查
# ===========================================================================
slide4 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(slide4)
add_header(slide4, "第一章 介面與導覽", "1.4 航運文書處理必備快速鍵速查清單", 4)
add_footer(slide4)

# 8 Shortcut Cards in 2x4 grid
sc_list = [
    ("Ctrl + E", "閃電填入 (Flash Fill)", "手打 1 個示範樣本後按 Ctrl+E，瞬間自動辨識規律並智慧拆解全欄船名或航次！", TEXT_CYAN),
    ("Ctrl + Shift + L", "啟用 / 取消自動篩選", "瞬間在標題列加上下拉按鈕，便於排查空值、重複值或異常提單號碼。", TEXT_WHITE),
    ("Ctrl + Alt + V", "選擇性貼上 (Paste Special)", "清洗完成後，將公式計算結果貼上為「值 (V)」，切斷公式相依轉換為純淨主檔。", TEXT_EMERALD),
    ("Ctrl + H", "尋找與取代 (Find & Replace)", "批次將文字中的點號 '.' 或破折號 '-' 全數替換為標準斜線 '/' 或消除多餘空格。", TEXT_AMBER),
    ("F4 鍵", "重複上一步 / 鎖定座標", "在公式編輯時快速切換絕對引用 ($A$1)；在格式調整時重複上個動作。", TEXT_CYAN),
    ("F2 鍵", "直接進入編輯模式", "游標直接跳入儲存格末端編輯文字，無須雙擊滑鼠，快速修改提單字串。", TEXT_WHITE),
    ("Ctrl + 方向鍵", "跳轉至資料極限端", "瞬間跳轉至清單第 1 列或最底端（如第 5,000 列），無須緩慢滾動滑鼠輪。", TEXT_EMERALD),
    ("Alt + Enter", "儲存格內強制換行", "在單一儲存格內將託運人名稱與地址分行排版，符合提單列印規範。", TEXT_AMBER)
]

for idx, (skey, sname, sdesc, scol) in enumerate(sc_list):
    col_i = idx % 4
    row_i = idx // 4
    sx = Inches(0.8 + col_i * 2.98)
    sy = Inches(1.6 + row_i * 2.6)
    
    scard = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sx, sy, Inches(2.85), Inches(2.4))
    scard.fill.solid()
    scard.fill.fore_color.rgb = CARD_BG
    scard.line.color.rgb = scol
    
    stf = scard.text_frame
    stf.word_wrap = True
    stf.margin_top = Inches(0.2)
    stf.margin_left = Inches(0.2)
    stf.margin_right = Inches(0.2)
    
    p0 = stf.paragraphs[0]
    r0 = p0.add_run()
    r0.text = skey
    r0.font.name = "Consolas"
    r0.font.size = Pt(13)
    r0.font.bold = True
    r0.font.color.rgb = scol
    
    p1 = stf.add_paragraph()
    p1.space_before = Pt(4)
    r1 = p1.add_run()
    r1.text = sname
    r1.font.name = "Microsoft JhengHei"
    r1.font.size = Pt(11)
    r1.font.bold = True
    r1.font.color.rgb = TEXT_WHITE
    
    p2 = stf.add_paragraph()
    p2.space_before = Pt(6)
    r2 = p2.add_run()
    r2.text = sdesc
    r2.font.name = "Microsoft JhengHei"
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = TEXT_MUTED

set_speaker_notes(slide4, "快速鍵能夠讓行政人員在鍵盤上維持高速節奏。特別強調 Ctrl+E 閃電填入與 Ctrl+Shift+L 快速篩選，這兩個是日常使用率最高的組合。")

# ===========================================================================
# SLIDE 5: 什麼是乾淨資料 (Tidy Data)？航運數據三大原則
# ===========================================================================
slide5 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(slide5)
add_header(slide5, "第二章 核心常識", "2.1 什麼是「乾淨資料」(Tidy Data)？航運數據三大原則", 5)
add_footer(slide5)

# Subtitle banner
tb_s5 = slide5.shapes.add_textbox(Inches(0.8), Inches(1.35), Inches(11.733), Inches(0.4))
tb_s5.text_frame.paragraphs[0].text = "資料科學大師 Hadley Wickham 提出的經典架構，符合此架構方能無障礙進行樞紐分析與統計"
tb_s5.text_frame.paragraphs[0].font.name = "Microsoft JhengHei"
tb_s5.text_frame.paragraphs[0].font.size = Pt(11.5)
tb_s5.text_frame.paragraphs[0].font.color.rgb = TEXT_MUTED

# 3 Principles Big Cards
p_cards = [
    ("原則一：每一列 (Row) 為單一實體",
     "• 每一列代表單一獨立的提單（B/L No.）或貨櫃流水紀錄。\n• 嚴禁將兩個貨櫃合併寫在同一列。\n• 嚴禁在資料中間插入空白分隔列。\n\n實務範例：\n[X] 提單列中填寫：EMCU123 / EMCU456\n[√] 拆分為兩列，各自對應獨立櫃號",
     TEXT_CYAN),
    ("原則二：每一欄 (Col) 為單一變數",
     "• 起運港 (POL)、目的港 (POD)、結關日、毛重各自成獨立欄。\n• 嚴禁將不同概念塞入同一欄位中。\n• 每一欄必須有唯一的欄位標題。\n\n實務範例：\n[X] 欄名：船名與航次 (合併)\n[√] 欄位A：船名  |  欄位B：航次",
     TEXT_EMERALD),
    ("原則三：每一格 (Cell) 為單一原子值",
     "• 儲存格僅包含無法再分割的最小純粹數值。\n• 數值必須是純數字，不可包含英文字母或單位。\n• 日期必須是標準西元序列值。\n\n實務範例：\n[X] 儲存格填寫：USD 2,450.00\n[√] 金額欄：2450.00  |  幣別欄：USD",
     TEXT_AMBER)
]

for idx, (ptitle, pdesc, pcol) in enumerate(p_cards):
    px = Inches(0.8 + idx * 4.0)
    pcard = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px, Inches(1.8), Inches(3.7), Inches(5.0))
    pcard.fill.solid()
    pcard.fill.fore_color.rgb = CARD_BG
    pcard.line.color.rgb = pcol
    pcard.line.width = Pt(1.5)
    
    ptf = pcard.text_frame
    ptf.word_wrap = True
    ptf.margin_top = Inches(0.25)
    ptf.margin_left = Inches(0.25)
    ptf.margin_right = Inches(0.25)
    
    p0 = ptf.paragraphs[0]
    r0 = p0.add_run()
    r0.text = ptitle
    r0.font.name = "Microsoft JhengHei"
    r0.font.size = Pt(13)
    r0.font.bold = True
    r0.font.color.rgb = pcol
    
    p1 = ptf.add_paragraph()
    p1.space_before = Pt(12)
    r1 = p1.add_run()
    r1.text = pdesc
    r1.font.name = "Microsoft JhengHei"
    r1.font.size = Pt(10.5)
    r1.font.color.rgb = TEXT_WHITE

set_speaker_notes(slide5, "強調原子化觀念。只有每個儲存格都是單一純粹的值，才能進行分組、計算平均與樞紐分析。")

# ===========================================================================
# SLIDE 6: 儲存格資料型態原理與實務陷阱
# ===========================================================================
slide6 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(slide6)
add_header(slide6, "第二章 核心常識", "2.2 儲存格資料型態原理與實務陷阱", 6)
add_footer(slide6)

# Left Column: 4 Types
tb_l6 = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(5.7), Inches(5.3))
tb_l6.fill.solid()
tb_l6.fill.fore_color.rgb = CARD_BG
tb_l6.line.color.rgb = CARD_BORDER
l_tf6 = tb_l6.text_frame
l_tf6.word_wrap = True
l_tf6.margin_top = Inches(0.2)
l_tf6.margin_left = Inches(0.25)

lp0 = l_tf6.paragraphs[0]
lr0 = lp0.add_run()
lr0.text = "Excel 四大基礎資料型態與特性"
lr0.font.name = "Microsoft JhengHei"
lr0.font.size = Pt(13)
lr0.font.bold = True
lr0.font.color.rgb = TEXT_CYAN

type_items = [
    ("● 文字 (Text) - 預設靠左對齊", "純字串組合，不可做四則運算。如提單號碼、託運人。數字若存成文字，SUM 加總將忽略顯示為 0。"),
    ("● 數值 (Number) - 預設靠右對齊", "純浮點數或整數，可加減乘除。夾雜英文字母或單位符號 (如 KGS) 會被強制降級為文字！"),
    ("● 日期時間 (Date) - 預設靠右對齊", "實質為自 1900/1/1 起算的連續序列整數。因此兩日期可以直接相減計算天數！"),
    ("● 邏輯值 (Boolean) - 預設居中對齊", "僅 TRUE (真) 或 FALSE (偽) 兩種狀態。手動打字之 'TRUE' 與邏輯值不同。")
]
for t_title, t_desc in type_items:
    p = l_tf6.add_paragraph()
    p.space_before = Pt(10)
    r_t = p.add_run()
    r_t.text = t_title + "\n"
    r_t.font.name = "Microsoft JhengHei"
    r_t.font.size = Pt(11)
    r_t.font.bold = True
    r_t.font.color.rgb = TEXT_WHITE
    r_d = p.add_run()
    r_d.text = t_desc
    r_d.font.name = "Microsoft JhengHei"
    r_d.font.size = Pt(10)
    r_d.font.color.rgb = TEXT_MUTED

# Right Column: Error Codes
tb_r6 = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.3))
tb_r6.fill.solid()
tb_r6.fill.fore_color.rgb = CARD_BG
tb_r6.line.color.rgb = CARD_BORDER
r_tf6 = tb_r6.text_frame
r_tf6.word_wrap = True
r_tf6.margin_top = Inches(0.2)
r_tf6.margin_left = Inches(0.25)

rp0 = r_tf6.paragraphs[0]
rr0 = rp0.add_run()
rr0.text = "航運常見錯誤代碼 (#) 原因與對策"
rr0.font.name = "Microsoft JhengHei"
rr0.font.size = Pt(13)
rr0.font.bold = True
rr0.font.color.rgb = TEXT_ROSE

error_items = [
    ("#N/A (查無資料 / 比對失敗)", "成因：提單號前後有多餘空白，VLOOKUP 找不到。\n解法：TRIM(A2) 去除空白，或包 IFERROR() 防呆。"),
    ("#VALUE! (型態不相容)", "成因：拿包含文字的 '18450 KGS' 去做運算乘法。\n解法：使用 SUBSTITUTE 消除單位轉純數值。"),
    ("#REF! (參照儲存格已遺失)", "成因：公式引用的欄列被直接整欄刪除。\n解法：立即按 Ctrl+Z 撤銷或重新指定參照座標。"),
    ("#DIV/0! (除數為 0 或空值)", "成因：計算平均每櫃運費時，總櫃數為 0 或空白。\n解法：使用 IF 判斷除數大於 0 再行計算。")
]
for e_title, e_desc in error_items:
    p = r_tf6.add_paragraph()
    p.space_before = Pt(10)
    r_t = p.add_run()
    r_t.text = e_title + "\n"
    r_t.font.name = "Consolas"
    r_t.font.size = Pt(11)
    r_t.font.bold = True
    r_t.font.color.rgb = TEXT_AMBER
    r_d = p.add_run()
    r_d.text = e_desc
    r_d.font.name = "Microsoft JhengHei"
    r_d.font.size = Pt(10)
    r_d.font.color.rgb = TEXT_MUTED

set_speaker_notes(slide6, "教導學員用眼睛判斷型態：文字靠左、數值靠右。重點說明日期在 Excel 底層是連續整數，因此才能直接相減算航期或滯期天數。")

# ===========================================================================
# SLIDE 7: 航運實務常見「七大髒資料症狀」
# ===========================================================================
slide7 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(slide7)
add_header(slide7, "第二章 核心常識", "2.3 航運實務常見「七大髒資料症狀」與衝擊剖析", 7)
add_footer(slide7)

# Left Image
img_path_3 = os.path.join(fig_dir, "fig3_clean_vs_dirty_data.png")
if os.path.exists(img_path_3):
    slide7.shapes.add_picture(img_path_3, Inches(0.8), Inches(1.5), Inches(6.5), Inches(5.2))

# Right 7 Symptoms List
tb_s7 = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.5), Inches(1.5), Inches(5.0), Inches(5.2))
tb_s7.fill.solid()
tb_s7.fill.fore_color.rgb = CARD_BG
tb_s7.line.color.rgb = CARD_BORDER
s7_tf = tb_s7.text_frame
s7_tf.word_wrap = True
s7_tf.margin_top = Inches(0.18)
s7_tf.margin_left = Inches(0.2)

s7_p0 = s7_tf.paragraphs[0]
s7_p0.text = "常見髒資料症狀與快速對策清單"
s7_p0.font.name = "Microsoft JhengHei"
s7_p0.font.size = Pt(12)
s7_p0.font.bold = True
s7_p0.font.color.rgb = TEXT_CYAN

sym_list = [
    ("症狀 1：多餘隱藏空格", "\"  EMCU20491823  \" → TRIM(A2)"),
    ("症狀 2：大小寫混亂不一", "\"twkhh\" / \"uslgb\" → UPPER(A2)"),
    ("症狀 3：複合欄位未原子化", "\"EVER GIVEN / 055W\" → 資料剖析 / Ctrl+E"),
    ("症狀 4：數值夾雜計量單位", "\"18,450 KGS\" → SUBSTITUTE + VALUE"),
    ("症狀 5：日期格式五花八門", "\"2026.11.05\" → DATE 統一西元標準日"),
    ("症狀 6：非標準代碼與俗稱", "\"40'HC\", \"40 DRY\" → XLOOKUP / 資料驗證"),
    ("症狀 7：重複提單重疊登錄", "同一提單多次出現 → 移除重複項")
]
for stitle, ssol in sym_list:
    p = s7_tf.add_paragraph()
    p.space_before = Pt(8)
    r0 = p.add_run()
    r0.text = stitle + "\n"
    r0.font.name = "Microsoft JhengHei"
    r0.font.size = Pt(10)
    r0.font.bold = True
    r0.font.color.rgb = TEXT_WHITE
    r1 = p.add_run()
    r1.text = ssol
    r1.font.name = "Consolas"
    r1.font.size = Pt(9.5)
    r1.font.color.rgb = TEXT_AMBER

set_speaker_notes(slide7, "讓學員對照圖2-2，檢視日常經手的報表是否也有這七大毛病。指出每一項毛病都有對應的解法。")

# ===========================================================================
# SLIDE 8: 資料前處理標準五步驟循環 SOP
# ===========================================================================
slide8 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(slide8)
add_header(slide8, "第二章 核心常識", "2.4 資料前處理標準五步驟循環 SOP 流程", 8)
add_footer(slide8)

# Center Image
img_path_2 = os.path.join(fig_dir, "fig2_data_preprocessing_flow.png")
if os.path.exists(img_path_2):
    slide8.shapes.add_picture(img_path_2, Inches(1.5), Inches(1.5), Inches(10.333), Inches(3.2))

# 5 Step Cards Bottom
sop_steps = [
    ("步驟 1", "結構檢視與備份", "保護原始資料不變更\n確認標題唯一無空白\n檢查資料型態分佈", TEXT_CYAN),
    ("步驟 2", "字串雜質清理", "TRIM 去除多餘空格\nUPPER 統一英文大寫\nCLEAN 清除不可見字元", TEXT_EMERALD),
    ("步驟 3", "欄位原子化分割", "資料剖析拆船名航次\n分離包裝件數與單位\n提取純數值與幣別", TEXT_AMBER),
    ("步驟 4", "型態規格化轉換", "統一西元 YYYY-MM-DD\n文字數字轉運算數值\n空值與錯誤值標記防呆", TEXT_WHITE),
    ("步驟 5", "去重比對與驗證", "移除重複提單號碼\nXLOOKUP 港口碼校對\n確認滿足分析要求就緒", TEXT_ROSE)
]
for idx, (snum, stitle, sdesc, scol) in enumerate(sop_steps):
    sx = Inches(0.8 + idx * 2.4)
    scard = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sx, Inches(4.9), Inches(2.25), Inches(1.9))
    scard.fill.solid()
    scard.fill.fore_color.rgb = CARD_BG
    scard.line.color.rgb = scol
    
    stf = scard.text_frame
    stf.word_wrap = True
    stf.margin_top = Inches(0.12)
    stf.margin_left = Inches(0.12)
    stf.margin_right = Inches(0.12)
    
    p0 = stf.paragraphs[0]
    p0.alignment = PP_ALIGN.CENTER
    r0 = p0.add_run()
    r0.text = f"{snum} : {stitle}"
    r0.font.name = "Microsoft JhengHei"
    r0.font.size = Pt(10)
    r0.font.bold = True
    r0.font.color.rgb = scol
    
    p1 = stf.add_paragraph()
    p1.space_before = Pt(4)
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run()
    r1.text = sdesc
    r1.font.name = "Microsoft JhengHei"
    r1.font.size = Pt(8.5)
    r1.font.color.rgb = TEXT_MUTED

set_speaker_notes(slide8, "前處理五步 SOP 是本手冊的方法論核心：先備份 -> 去雜質 -> 原子化 -> 規格化轉型 -> 去重驗收。養成標準SOP就不會遺漏。")

# ===========================================================================
# SLIDE 9: 30筆航運示範案例：原始結構與問題診斷
# ===========================================================================
slide9 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(slide9)
add_header(slide9, "第三章 示範案例", "3.1 30筆航運示範案例：原始資料結構與問題診斷", 9)
add_footer(slide9)

# Subtitle
tb_s9 = slide9.shapes.add_textbox(Inches(0.8), Inches(1.35), Inches(11.733), Inches(0.4))
tb_s9.text_frame.paragraphs[0].text = "刻意設計包含 7 大常見髒資料特徵的示範清單，模擬從舊系統或外部匯出的真實瑕疵報表"
tb_s9.text_frame.paragraphs[0].font.name = "Microsoft JhengHei"
tb_s9.text_frame.paragraphs[0].font.size = Pt(11)
tb_s9.text_frame.paragraphs[0].font.color.rgb = TEXT_MUTED

# Table Box
t_rows = 6
t_cols = 6
tbl_shape = slide9.shapes.add_table(t_rows, t_cols, Inches(0.8), Inches(1.8), Inches(11.733), Inches(3.2))
tbl = tbl_shape.table

# Set Column Widths
tbl.columns[0].width = Inches(0.8)
tbl.columns[1].width = Inches(2.2)
tbl.columns[2].width = Inches(2.4)
tbl.columns[3].width = Inches(1.8)
tbl.columns[4].width = Inches(1.8)
tbl.columns[5].width = Inches(2.733)

tbl_headers = ["項次", "提單號碼 (髒)", "船名航次 (複合)", "毛重 (文字)", "結關日 (混雜)", "瑕疵診斷與衝擊"]
for c_i, th in enumerate(tbl_headers):
    cell = tbl.cell(0, c_i)
    cell.fill.solid()
    cell.fill.fore_color.rgb = RGBColor(27, 54, 93)
    p = cell.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = th
    r.font.name = "Microsoft JhengHei"
    r.font.size = Pt(9.5)
    r.font.bold = True
    r.font.color.rgb = TEXT_WHITE

sample_tbl_data = [
    ["1", "  EMCU20491823 ", "EVER GIVEN / 0124-055W", "18,450.50 KGS", "2026/11/02", "首尾空格、複合船名、重量含文字"],
    ["2", "OOCLU918273645", "OOCL ROTTERDAM / 088E", "12400", "2026-11-03", "港口小寫 (tw tpe)、減號日期"],
    ["3", "MSCU992817263", "MSC ISABELLA / 2311W", "9,800.00 KGS", "2026.11.04", "點號非標準日期、運費含$符號"],
    ["4", "YMLU019283746", "YM WELLHEAD / 221E", "5600 KGS", "2026/11/05", "★ 與項次 28 整筆完全重複！"],
    ["28", "YMLU019283746", "YM WELLHEAD / 221E", "5600 KGS", "2026/11/05", "★ 重複登錄無效垃圾資料，虛增貨量"]
]
for r_i, r_data in enumerate(sample_tbl_data):
    for c_i, val in enumerate(r_data):
        cell = tbl.cell(r_i + 1, c_i)
        cell.fill.solid()
        bg_c = RGBColor(40, 25, 25) if "重複" in val or "★" in val else (RGBColor(15, 25, 45) if r_i % 2 == 1 else RGBColor(12, 20, 38))
        cell.fill.fore_color.rgb = bg_c
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER if c_i in [0, 4] else PP_ALIGN.LEFT
        r = p.add_run()
        r.text = val
        r.font.name = "Consolas" if c_i in [1, 2, 3, 4] else "Microsoft JhengHei"
        r.font.size = Pt(9)
        r.font.color.rgb = TEXT_AMBER if "★" in val else TEXT_WHITE

# Bottom Tip
b_box9 = slide9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.3), Inches(11.733), Inches(1.5))
b_box9.fill.solid()
b_box9.fill.fore_color.rgb = CARD_BG
b_box9.line.color.rgb = CARD_BORDER
b_tf9 = b_box9.text_frame
b_tf9.margin_top = Inches(0.15)
b_tf9.margin_left = Inches(0.25)
b_p9 = b_tf9.paragraphs[0]
b_p9.text = "📂 隨附活頁簿對應：c:\\proj\\w12\\航運行政Excel資料前處理_示範資料檔_30筆.xlsx\n"
b_p9.font.name = "Microsoft JhengHei"
b_p9.font.size = Pt(10)
b_p9.font.bold = True
b_p9.font.color.rgb = TEXT_CYAN
b_p9_1 = b_tf9.add_paragraph()
b_p9_1.space_before = Pt(3)
b_p9_1.text = "內含【01_原始待處理資料(Dirty)】與【02_前處理完成資料(Cleaned)】工作表，學員可直接打開活頁簿對照演練。"
b_p9_1.font.name = "Microsoft JhengHei"
b_p9_1.font.size = Pt(9.5)
b_p9_1.font.color.rgb = TEXT_WHITE

set_speaker_notes(slide9, "引導學員觀摩表3-1。說明項次4與28完全重複，提單夾帶空格、重量有KGS文字。這就是即將在下一頁逐一清洗的目標。")

# ===========================================================================
# SLIDE 10: 30筆示範案例：清洗流程實戰八大步驟
# ===========================================================================
slide10 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(slide10)
add_header(slide10, "第三章 示範案例", "3.3~3.10 示範案例：清洗流程實戰八大步驟指南", 10)
add_footer(slide10)

step8_list = [
    ("步驟 1：備份工作表", "右鍵工作表標籤選【移動或複製】-> 勾選【建立複本】，命名為『清洗操作工作表』。"),
    ("步驟 2：提單去空格大寫", "=UPPER(TRIM(B2))\n消除首尾空格並強制全大寫，雙擊向下填滿後「貼上值」。"),
    ("步驟 3：移除重複提單", "點選【資料】->【移除重複項】，僅勾選【提單號碼】，成功剔除項次 28，保留 29 筆唯一值。"),
    ("步驟 4：船名航次剖析", "使用【資料剖析】依斜線 '/' 切分，或手打一個範例後按快速鍵 Ctrl+E 閃電填入！"),
    ("步驟 5：港口碼標準化", "=UPPER(TRIM(D2))\n清除代碼內部空格，統一轉為 5 碼聯合國標準 UN/LOCODE。"),
    ("步驟 6：毛重去除單位轉值", "=VALUE(SUBSTITUTE(SUBSTITUTE(I2,\"KGS\",\"\"),\",\",\"\"))\n去除文字單位與逗點轉純數值。"),
    ("步驟 7：統一標準西元日", "Ctrl+H 將點號 '.' 取代為斜線 '/'，儲存格格式設定為 YYYY-MM-DD。"),
    ("步驟 8：運費純數值化", "=VALUE(SUBSTITUTE(SUBSTITUTE(N2,\"USD\",\"\"),\"$\",\"\"))\n消除貨幣符號轉浮點數。")
]
for idx, (stitle, sdesc) in enumerate(step8_list):
    col_i = idx % 4
    row_i = idx // 4
    sx = Inches(0.8 + col_i * 2.98)
    sy = Inches(1.6 + row_i * 2.6)
    
    scard = slide10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sx, sy, Inches(2.85), Inches(2.4))
    scard.fill.solid()
    scard.fill.fore_color.rgb = CARD_BG
    scard.line.color.rgb = TEXT_CYAN if idx in [1, 3, 5] else CARD_BORDER
    
    stf = scard.text_frame
    stf.word_wrap = True
    stf.margin_top = Inches(0.18)
    stf.margin_left = Inches(0.2)
    stf.margin_right = Inches(0.2)
    
    p0 = stf.paragraphs[0]
    r0 = p0.add_run()
    r0.text = stitle
    r0.font.name = "Microsoft JhengHei"
    r0.font.size = Pt(11)
    r0.font.bold = True
    r0.font.color.rgb = TEXT_CYAN
    
    p1 = stf.add_paragraph()
    p1.space_before = Pt(8)
    r1 = p1.add_run()
    r1.text = sdesc
    r1.font.name = "Consolas" if "=" in sdesc else "Microsoft JhengHei"
    r1.font.size = Pt(9.5)
    r1.font.color.rgb = TEXT_WHITE

set_speaker_notes(slide10, "逐步解說清洗八步驟。步驟2、4、6是三大核心技巧：TRIM+UPPER標準化主鍵、資料剖析原子化拆分、SUBSTITUTE+VALUE去除單位轉為可運算數值。")

# ===========================================================================
# SLIDE 11: 30筆示範案例：清洗前後成果對比展示
# ===========================================================================
slide11 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(slide11)
add_header(slide11, "第三章 示範案例", "3.11 30筆示範案例：清洗前 (Dirty) 與 清洗後 (Clean) 成效對照", 11)
add_footer(slide11)

# Left Image Box
img_path_9 = os.path.join(fig_dir, "fig9_demo30_before_after.png")
if os.path.exists(img_path_9):
    slide11.shapes.add_picture(img_path_9, Inches(0.8), Inches(1.5), Inches(7.2), Inches(4.5))

# Right Comparison Text Box
tb_cmp = slide11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.3), Inches(1.5), Inches(4.2), Inches(5.2))
tb_cmp.fill.solid()
tb_cmp.fill.fore_color.rgb = CARD_BG
tb_cmp.line.color.rgb = CARD_BORDER
ctf = tb_cmp.text_frame
ctf.word_wrap = True
ctf.margin_top = Inches(0.2)
ctf.margin_left = Inches(0.25)
ctf.margin_right = Inches(0.25)

p_c0 = ctf.paragraphs[0]
p_c0.text = "[X] 清洗前原始狀態 (Dirty)"
p_c0.font.name = "Microsoft JhengHei"
p_c0.font.size = Pt(11)
p_c0.font.bold = True
p_c0.font.color.rgb = TEXT_ROSE

p_c1 = ctf.add_paragraph()
p_c1.space_before = Pt(4)
p_c1.text = "• 提單夾帶空格，比對拋出 #N/A 錯誤\n• 毛重欄位加總 SUM 結果為 0.00\n• 日期混雜無法計算在場與航行天數\n• 重複登錄 1 筆無效紀錄，虛增貨量"
p_c1.font.name = "Microsoft JhengHei"
p_c1.font.size = Pt(9.5)
p_c1.font.color.rgb = TEXT_MUTED

p_c2 = ctf.add_paragraph()
p_c2.space_before = Pt(14)
p_c2.text = "[√] 清洗後乾淨主檔 (Cleaned)"
p_c2.font.name = "Microsoft JhengHei"
p_c2.font.size = Pt(11)
p_c2.font.bold = True
p_c2.font.color.rgb = TEXT_EMERALD

p_c3 = ctf.add_paragraph()
p_c3.space_before = Pt(4)
p_c3.text = "• 提單純淨大寫，XLOOKUP 比對成功率 100%\n• 毛重加總正確得出總貨重 484,700 KGS\n• 船名與航次原子化獨立兩欄\n• 精確保留 29 筆唯一真實營收主檔"
p_c3.font.name = "Microsoft JhengHei"
p_c3.font.size = Pt(9.5)
p_c3.font.color.rgb = TEXT_WHITE

# Bottom Metric Callout
b_box11 = slide11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.15), Inches(7.2), Inches(0.8))
b_box11.fill.solid()
b_box11.fill.fore_color.rgb = RGBColor(12, 35, 60)
b_box11.line.color.rgb = TEXT_CYAN
b_p11 = b_box11.text_frame.paragraphs[0]
b_p11.alignment = PP_ALIGN.CENTER
b_r11 = b_p11.add_run()
b_r11.text = "清洗後成果驗收：總筆數 = 29 筆  |  總毛重 = 484,700 KGS  |  總運費 = USD 64,550.00"
b_r11.font.name = "Microsoft JhengHei"
b_r11.font.size = Pt(10.5)
b_r11.font.bold = True
b_r11.font.color.rgb = TEXT_CYAN

set_speaker_notes(slide11, "清洗前後成效對照。指出原本毛重加總是0，清洗後成功轉為浮點數，總重量484,700公斤，總運費64,550美元，去除了重複項次28。")

# ===========================================================================
# SLIDE 12: 前處理四大必備選單功能
# ===========================================================================
slide12 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(slide12)
add_header(slide12, "第四章 功能與函數", "4.1 功能區選單操作手冊：四大必備前處理工具深入解析", 12)
add_footer(slide12)

tools_4 = [
    ("1. 資料剖析 (Text to Cols)", "fig4_text_to_columns_wizard.png", "依斜線、空格或破折號等分隔符號，精準將複合欄位切分至相鄰欄。", TEXT_CYAN),
    ("2. 移除重複項 (Duplicates)", "fig5_remove_duplicates_dialog.png", "防呆秘訣：絕不可全選所有欄位，僅勾選「提單號」或「櫃號」唯一主鍵。", TEXT_ROSE),
    ("3. 快速填入 (Flash Fill)", "fig6_flash_fill_demo.png", "智慧模式學習：手動鍵入 1 個示範後按 Ctrl+E，自動完成全欄字串拆解。", TEXT_AMBER),
    ("4. 資料驗證 (Validation)", "fig7_data_validation_dialog.png", "建立儲存格下拉選單（限定 20GP, 40GP, 40HQ），徹底防呆阻絕錯誤輸入。", TEXT_EMERALD)
]

for idx, (t_name, t_img, t_desc, t_col) in enumerate(tools_4):
    tx = Inches(0.8 + idx * 2.98)
    tcard = slide12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, tx, Inches(1.5), Inches(2.85), Inches(5.3))
    tcard.fill.solid()
    tcard.fill.fore_color.rgb = CARD_BG
    tcard.line.color.rgb = t_col
    
    ttf = tcard.text_frame
    ttf.word_wrap = True
    ttf.margin_top = Inches(0.15)
    ttf.margin_left = Inches(0.18)
    ttf.margin_right = Inches(0.18)
    
    p0 = ttf.paragraphs[0]
    r0 = p0.add_run()
    r0.text = t_name
    r0.font.name = "Microsoft JhengHei"
    r0.font.size = Pt(11)
    r0.font.bold = True
    r0.font.color.rgb = t_col
    
    # Add thumbnail picture
    img_t_path = os.path.join(fig_dir, t_img)
    if os.path.exists(img_t_path):
        slide12.shapes.add_picture(img_t_path, tx + Inches(0.15), Inches(2.1), Inches(2.55), Inches(2.3))
        
    p1 = ttf.add_paragraph()
    p1.space_before = Pt(185) # leave space for image
    r1 = p1.add_run()
    r1.text = t_desc
    r1.font.name = "Microsoft JhengHei"
    r1.font.size = Pt(9.5)
    r1.font.color.rgb = TEXT_WHITE

set_speaker_notes(slide12, "介紹功能區四大工具。特別強調「移除重複項」時不可全選欄位，要單獨勾選提單號；「資料驗證」可以從源頭防止輸入錯誤櫃型。")

# ===========================================================================
# SLIDE 13: 航運前處理「十大核心函數」體系架構導圖
# ===========================================================================
slide13 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(slide13)
add_header(slide13, "第四章 功能與函數", "4.2 航運前處理「十大核心函數」體系架構導圖", 13)
add_footer(slide13)

# Left Image
img_path_8 = os.path.join(fig_dir, "fig8_core_functions_architecture.png")
if os.path.exists(img_path_8):
    slide13.shapes.add_picture(img_path_8, Inches(0.8), Inches(1.5), Inches(7.2), Inches(4.5))

# Right 4 Categories Box
tb_f13 = slide13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.3), Inches(1.5), Inches(4.2), Inches(5.2))
tb_f13.fill.solid()
tb_f13.fill.fore_color.rgb = CARD_BG
tb_f13.line.color.rgb = CARD_BORDER
ftf13 = tb_f13.text_frame
ftf13.word_wrap = True
ftf13.margin_top = Inches(0.2)
ftf13.margin_left = Inches(0.22)
ftf13.margin_right = Inches(0.22)

f_cats = [
    ("1. 字串清洗類 (Text Cleaning)", "• TRIM(A2) : 清除首尾及多餘空格\n• CLEAN(A2) : 清除隱藏換行字元\n• UPPER(A2) : 英文統一強制大寫", TEXT_EMERALD),
    ("2. 字串截取替換 (Extraction)", "• LEFT / MID / RIGHT : 截取指定碼數\n• FIND(text, within) : 尋找符號位置\n• SUBSTITUTE(A2, old, new) : 精確替換", TEXT_CYAN),
    ("3. 型態數值轉換 (Type Casting)", "• VALUE(text) : 文字數字轉可運算數值\n• TEXT(num, fmt) : 格式化數值或日期", TEXT_AMBER),
    ("4. 日期重構與查詢防呆 (Date & Lookup)", "• DATE(Y, M, D) : 重構西元標準日期\n• XLOOKUP(key, lk, ret) : 跨表查找\n• IFERROR(val, err_val) : 攔截屏蔽錯誤", TEXT_WHITE)
]
for idx, (fc_title, fc_desc, fc_col) in enumerate(f_cats):
    p = ftf13.paragraphs[0] if idx == 0 else ftf13.add_paragraph()
    if idx > 0: p.space_before = Pt(8)
    r_t = p.add_run()
    r_t.text = fc_title + "\n"
    r_t.font.name = "Microsoft JhengHei"
    r_t.font.size = Pt(10.5)
    r_t.font.bold = True
    r_t.font.color.rgb = fc_col
    r_d = p.add_run()
    r_d.text = fc_desc
    r_d.font.name = "Consolas"
    r_d.font.size = Pt(9)
    r_d.font.color.rgb = TEXT_MUTED

# Bottom Tip
b_box13 = slide13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.15), Inches(7.2), Inches(0.8))
b_box13.fill.solid()
b_box13.fill.fore_color.rgb = RGBColor(12, 35, 60)
b_box13.line.color.rgb = TEXT_CYAN
b_p13 = b_box13.text_frame.paragraphs[0]
b_p13.text = "💡 實務精髓：將多個函數巢狀組合。例如 =VALUE(SUBSTITUTE(SUBSTITUTE(I2, \"KGS\", \"\"), \",\", \"\"))，一次完成去單位與轉純數值。"
b_p13.font.name = "Microsoft JhengHei"
b_p13.font.size = Pt(9.5)
b_p13.font.color.rgb = TEXT_WHITE

set_speaker_notes(slide13, "掌握四大體系十大函數：TRIM/CLEAN去空格、LEFT/MID/FIND截取、VALUE轉數值、XLOOKUP跨表查找。巢狀公式如同積木組合。")

# ===========================================================================
# SLIDE 14: 300筆航運實戰題庫四大業務模組全覽
# ===========================================================================
slide14 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(slide14)
add_header(slide14, "第五章 300筆題庫", "5.1~5.5 300筆實戰練習題庫四大業務模組全覽", 14)
add_footer(slide14)

mod_4 = [
    ("模組 A (項次 1 - 100)", "海運出口訂艙裝船清單清理", 
     "• 提單號去空格與連字號\n• 貨櫃 11 碼連字號去除\n• 件數與包裝單位拆分\n• 毛重純數值萃取可加總\n\n核心公式：\n=VALUE(SUBSTITUTE(SUBSTITUTE(I2,\"KGS\",\"\"),\",\",\"\"))", 
     TEXT_CYAN),
    ("模組 B (項次 101 - 170)", "貨櫃滯期延滯費 (D&D) 前處理", 
     "• 卸船日與還櫃日相減\n• IF(ISBLANK()) 未還櫃算至今天\n• MAX(0, 天數-免費期) 避免負數\n• 依櫃型計算每日滯期費\n\n核心公式：\n=MAX(0, IF(ISBLANK(M2), TODAY()-L2, M2-L2) - 7)", 
     TEXT_AMBER),
    ("模組 C (項次 171 - 240)", "貨代應收應付帳單幣別清理", 
     "• 外幣幣別代碼 (USD/EUR) 分離\n• 原幣純數值轉型\n• 依基準匯率 ROUND 折合 TWD\n• 5% 營業稅額與含稅總額\n\n核心公式：\n=ROUND(原幣純金額 * 折合匯率, 0)", 
     TEXT_EMERALD),
    ("模組 D (項次 241 - 300)", "跨國港口與船期表標準化", 
     "• 港口代碼去空白轉大寫\n• XLOOKUP 比對 UN/LOCODE 主檔\n• 自動對照中文港口站名\n• 計算 ETD-ETA 航程天數\n\n核心公式：\n=XLOOKUP(G2, 港口表!A:A, 港口表!B:B)", 
     TEXT_ROSE)
]

for idx, (m_code, m_title, m_desc, m_col) in enumerate(mod_4):
    mx = Inches(0.8 + idx * 2.98)
    mcard = slide14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, mx, Inches(1.5), Inches(2.85), Inches(5.3))
    mcard.fill.solid()
    mcard.fill.fore_color.rgb = CARD_BG
    mcard.line.color.rgb = m_col
    
    mtf = mcard.text_frame
    mtf.word_wrap = True
    mtf.margin_top = Inches(0.18)
    mtf.margin_left = Inches(0.2)
    mtf.margin_right = Inches(0.2)
    
    p0 = mtf.paragraphs[0]
    r0 = p0.add_run()
    r0.text = m_code
    r0.font.name = "Microsoft JhengHei"
    r0.font.size = Pt(10)
    r0.font.bold = True
    r0.font.color.rgb = m_col
    
    p1 = mtf.add_paragraph()
    p1.space_before = Pt(3)
    r1 = p1.add_run()
    r1.text = m_title
    r1.font.name = "Microsoft JhengHei"
    r1.font.size = Pt(11)
    r1.font.bold = True
    r1.font.color.rgb = TEXT_WHITE
    
    p2 = mtf.add_paragraph()
    p2.space_before = Pt(8)
    r2 = p2.add_run()
    r2.text = m_desc
    r2.font.name = "Consolas" if "=" in m_desc else "Microsoft JhengHei"
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = TEXT_MUTED

set_speaker_notes(slide14, "詳細解說隨附的300筆實戰題庫。包含四大情境：出口裝船清單、D&D延滯費、貨代帳單換算、跨國船期港口代碼。每模組皆有對應解題技巧。")

# ===========================================================================
# SLIDE 15: 航運資料品質控管三階段 SOP 檢核表
# ===========================================================================
slide15 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(slide15)
add_header(slide15, "第六章 品質控管", "6.1 航運行政人員資料前處理三階段 SOP 檢核表", 15)
add_footer(slide15)

sop_3stage = [
    ("第一階段：清洗前 (Pre-check)", 
     "• 原始檔案備份存檔：另存新檔並建立工作表複本，原始底稿永不變更。\n• 標題列唯一性檢查：確認第 1 列無合併儲存格、欄名唯一。\n• 初始總筆數記錄：選取資料欄記錄原始列數（如 300 列）。\n\n檢核目標：\n確保原始資料底稿安全，隨時可復原對照。",
     TEXT_CYAN),
    ("第二階段：清洗中 (In-process)", 
     "• 關鍵主鍵去空格大寫：提單號、櫃號套用 TRIM 與 UPPER。\n• 欄位原子化分割檢查：船名航次、件數單位必須各自獨立一欄。\n• 數值與日期真偽校驗：以 ISNUMBER 檢查重量與金額是否為真數值。\n\n檢核目標：\n確保數據無死角達成原子化與型態純淨。",
     TEXT_EMERALD),
    ("第三階段：清洗後 (Post-check)", 
     "• 重複記錄剔除確認：執行移除重複項，核對剔除筆數與原因。\n• 公式轉純值 (貼上值)：複製 -> 選擇性貼上為『值』，解除公式負擔。\n• 業務邏輯合理性審查：排查毛重是否負數、開航日是否早於結關日。\n\n檢核目標：\n確保分析資料 100% 合規並具備商業說服力。",
     TEXT_AMBER)
]

for idx, (stitle, sdesc, scol) in enumerate(sop_3stage):
    sx = Inches(0.8 + idx * 4.0)
    scard = slide15.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sx, Inches(1.6), Inches(3.7), Inches(5.2))
    scard.fill.solid()
    scard.fill.fore_color.rgb = CARD_BG
    scard.line.color.rgb = scol
    scard.line.width = Pt(1.5)
    
    stf = scard.text_frame
    stf.word_wrap = True
    stf.margin_top = Inches(0.2)
    stf.margin_left = Inches(0.25)
    stf.margin_right = Inches(0.25)
    
    p0 = stf.paragraphs[0]
    r0 = p0.add_run()
    r0.text = stitle
    r0.font.name = "Microsoft JhengHei"
    r0.font.size = Pt(12)
    r0.font.bold = True
    r0.font.color.rgb = scol
    
    p1 = stf.add_paragraph()
    p1.space_before = Pt(12)
    r1 = p1.add_run()
    r1.text = sdesc
    r1.font.name = "Microsoft JhengHei"
    r1.font.size = Pt(10.5)
    r1.font.color.rgb = TEXT_WHITE

set_speaker_notes(slide15, "品質控管是航運從業人員的核心素養。嚴格落實清洗前備份、清洗中驗證真偽、清洗後貼上值與邏輯審查，徹底避免粗心釀成大禍。")

# ===========================================================================
# SLIDE 16: 總結：隨附學習資源、活頁簿清單與國際參考文獻
# ===========================================================================
slide16 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(slide16)
add_header(slide16, "手冊總結", "總結：隨附學習資源、活頁簿清單與國際參考文獻", 16)
add_footer(slide16)

# Left Box: Files
tb_l16 = slide16.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(5.7), Inches(5.2))
tb_l16.fill.solid()
tb_l16.fill.fore_color.rgb = CARD_BG
tb_l16.line.color.rgb = CARD_BORDER
ltf16 = tb_l16.text_frame
ltf16.word_wrap = True
ltf16.margin_top = Inches(0.2)
ltf16.margin_left = Inches(0.25)

lp0 = ltf16.paragraphs[0]
lp0.text = "隨附正式活頁簿與教材電子檔"
lp0.font.name = "Microsoft JhengHei"
lp0.font.size = Pt(12)
lp0.font.bold = True
lp0.font.color.rgb = TEXT_CYAN

files_info = [
    ("1. 《航運行政Excel資料前處理_示範資料檔_30筆.xlsx》", "內含【01_原始待處理(Dirty)】與【02_前處理完成(Cleaned)】對照工作表。"),
    ("2. 《航運行政Excel資料前處理_綜合練習題庫_300筆.xlsx》", "收錄 300 筆四大情境題庫、全球港口主檔對照表及全套函數解答指引。"),
    ("3. 《航運行政人員_MS_Excel基礎資料前處理實戰手冊.docx》", "完整 60 頁詳細圖文教學手冊（含封面、目錄、圖表與頁碼編排）。"),
    ("4. 《航運行政_Excel資料前處理_簡報投影片.html》", "互動式 PPT 網頁版簡報（含內建公式實驗室與全螢幕投影功能）。")
]
for f_t, f_d in files_info:
    p = ltf16.add_paragraph()
    p.space_before = Pt(8)
    r_t = p.add_run()
    r_t.text = f_t + "\n"
    r_t.font.name = "Microsoft JhengHei"
    r_t.font.size = Pt(10)
    r_t.font.bold = True
    r_t.font.color.rgb = TEXT_WHITE
    r_d = p.add_run()
    r_d.text = f_d
    r_d.font.name = "Microsoft JhengHei"
    r_d.font.size = Pt(9.5)
    r_d.font.color.rgb = TEXT_MUTED

# Right Box: References
tb_r16 = slide16.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.2))
tb_r16.fill.solid()
tb_r16.fill.fore_color.rgb = CARD_BG
tb_r16.line.color.rgb = CARD_BORDER
rtf16 = tb_r16.text_frame
rtf16.word_wrap = True
rtf16.margin_top = Inches(0.2)
rtf16.margin_left = Inches(0.25)

rp0 = rtf16.paragraphs[0]
rp0.text = "國際權威參考文獻與引用來源"
rp0.font.name = "Microsoft JhengHei"
rp0.font.size = Pt(12)
rp0.font.bold = True
rp0.font.color.rgb = TEXT_EMERALD

refs_info = [
    ("● Microsoft Learn Documentation (2024)", "Excel Functions by Category (Text, Lookup, Date/Time). 參照官方函數定義與運算規範。"),
    ("● UN/CEFACT (UNECE) Recommendation No. 16", "UN/LOCODE: Codes for Trade and Transport Locations. 聯合國全球港口 5 碼代碼國際標準。"),
    ("● ISO 8601:2019 日期時間表示法標準", "規定國際標準西元日期表示法 (YYYY-MM-DD)，消除跨國日期格式分歧。"),
    ("● ISO 6346:2022 貨櫃編碼與識別標準", "規定貨櫃 11 碼箱號（4碼字頭+6碼流水號+1碼檢核碼）及尺寸型號代碼標準。"),
    ("● Hadley Wickham (2014) Tidy Data 論文", "現代資料科學乾淨資料結構設計原理之理論基礎。"),
    ("● FIATA (2023) 多式聯運作業指南", "海運提單（FBL）、訂艙確認件與國際承攬運費結算規範。")
]
for rf_t, rf_d in refs_info:
    p = rtf16.add_paragraph()
    p.space_before = Pt(6)
    r_t = p.add_run()
    r_t.text = rf_t + " : "
    r_t.font.name = "Microsoft JhengHei"
    r_t.font.size = Pt(9)
    r_t.font.bold = True
    r_t.font.color.rgb = TEXT_WHITE
    r_d = p.add_run()
    r_d.text = rf_d
    r_d.font.name = "Microsoft JhengHei"
    r_d.font.size = Pt(8.5)
    r_d.font.color.rgb = TEXT_MUTED

set_speaker_notes(slide16, "總結全文。鼓勵學員立即動手打開活頁簿練習，將學到的技能應用到日常海運行政工作中。註明各項國際規範參考來源。")

# Save Presentation
out_pptx_path = "c:\\proj\\w12\\航運行政人員_MS_Excel基礎資料前處理實戰指南.pptx"
prs.save(out_pptx_path)
print(f"PowerPoint presentation successfully saved to: {out_pptx_path}")
