import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np

# Ensure Chinese fonts work properly
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'Microsoft YaHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

out_dir = "c:\\proj\\w12\\figures"
os.makedirs(out_dir, exist_ok=True)

# -------------------------------------------------------------
# Figure 1: Excel 365 視窗工作介面全景導覽 (圖 1-1)
# -------------------------------------------------------------
def draw_fig1():
    fig, ax = plt.subplots(figsize=(11, 6.8), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    # Background window
    window = patches.FancyBboxPatch((1, 1), 98, 98, boxstyle="round,pad=0.2", fc="#FFFFFF", ec="#217346", lw=2)
    ax.add_patch(window)
    
    # Title bar (Green)
    title_bar = patches.Rectangle((1, 93), 98, 6, fc="#107C41", ec="none")
    ax.add_patch(title_bar)
    ax.text(3, 95.5, "Autosave [ON]   航運出口裝船總表_2026.xlsx - Excel", color="white", fontsize=11, fontweight='bold', va='center')
    ax.text(92, 95.5, "[ _ ] [ 口 ] [ X ]", color="white", fontsize=10, va='center')
    
    # Ribbon Tabs bar
    tabs_bar = patches.Rectangle((1, 88), 98, 5, fc="#F3F2F1", ec="#E1DFDD")
    ax.add_patch(tabs_bar)
    tabs = ["檔案 (File)", "常用 (Home)", "插入 (Insert)", "版面配置 (Layout)", "公式 (Formulas)", "資料 (Data)", "校閱 (Review)", "檢視 (View)"]
    x_pos = 3
    for i, tab in enumerate(tabs):
        if "資料" in tab:
            # Active tab highlight
            t_box = patches.Rectangle((x_pos-1, 88), 12, 5, fc="#FFFFFF", ec="#E1DFDD", lw=1)
            ax.add_patch(t_box)
            # Bottom green line for active tab
            ax.plot([x_pos-1, x_pos+11], [88.2, 88.2], color="#107C41", lw=3)
            ax.text(x_pos+5, 90.5, tab, color="#107C41", fontsize=9.5, fontweight='bold', ha='center', va='center')
            x_pos += 13
        else:
            ax.text(x_pos+4, 90.5, tab, color="#323130", fontsize=9.5, ha='center', va='center')
            x_pos += 11.5
            
    # Ribbon Toolbar icons area (Data tab)
    ribbon = patches.Rectangle((1, 72), 98, 16, fc="#FFFFFF", ec="#EDEBE9")
    ax.add_patch(ribbon)
    
    # Ribbon groups
    # Group 1: 取得並轉換資料
    ax.text(7, 83, "[取得資料]", fontsize=8.5, ha='center', color="#201F1E")
    ax.text(7, 78, "[文字/CSV]", fontsize=8.5, ha='center', color="#201F1E")
    ax.text(7, 73.5, "取得資料", fontsize=8, color="#605E5C", ha='center')
    ax.axvline(14, 0.73, 0.87, color="#E1DFDD", lw=1)
    
    # Group 2: 查詢與連線
    ax.text(20, 80, "[重新整理]", fontsize=8.5, ha='center', color="#201F1E")
    ax.text(20, 73.5, "查詢與連線", fontsize=8, color="#605E5C", ha='center')
    ax.axvline(26, 0.73, 0.87, color="#E1DFDD", lw=1)
    
    # Group 3: 排序與篩選
    ax.text(32, 83, "[ A-Z 遞增 ]", fontsize=8.5, ha='center', color="#201F1E")
    ax.text(32, 78, "[ 篩選 Filter ]", fontsize=8.5, ha='center', color="#107C41", fontweight='bold')
    ax.text(32, 73.5, "排序與篩選", fontsize=8, color="#605E5C", ha='center')
    ax.axvline(38, 0.73, 0.87, color="#E1DFDD", lw=1)
    
    # Group 4: 資料工具 (HIGHLIGHTED FOR PREPROCESSING!)
    tool_box = patches.FancyBboxPatch((39, 72.5), 37, 15, boxstyle="round,pad=0.2", fc="#EFF6FC", ec="#0078D4", lw=1.5, ls="--")
    ax.add_patch(tool_box)
    ax.text(45, 83, "[資料剖析 Text to Cols]", fontsize=8.5, ha='center', color="#004578", fontweight='bold')
    ax.text(45, 78, "[閃電填入 Flash Fill]", fontsize=8.5, ha='center', color="#004578", fontweight='bold')
    ax.text(61, 83, "[移除重複項 Duplicates]", fontsize=8.5, ha='center', color="#A80000", fontweight='bold')
    ax.text(61, 78, "[資料驗證 Validation]", fontsize=8.5, ha='center', color="#004578", fontweight='bold')
    ax.text(57, 74, "★ 資料前處理核心工具群 (Data Tools) ★", fontsize=8.5, color="#0078D4", ha='center', fontweight='bold')
    ax.axvline(77, 0.73, 0.87, color="#E1DFDD", lw=1)
    
    # Group 5: 預測與大綱
    ax.text(85, 80, "[分組大綱]", fontsize=8.5, ha='center', color="#201F1E")
    ax.text(85, 73.5, "大綱群組", fontsize=8, color="#605E5C", ha='center')
    
    # Name Box & Formula Bar area
    f_bar_bg = patches.Rectangle((1, 65), 98, 7, fc="#FAF9F8", ec="#EDEBE9")
    ax.add_patch(f_bar_bg)
    
    # Name Box
    name_box = patches.Rectangle((3, 66.5), 10, 4.2, fc="#FFFFFF", ec="#8A8886")
    ax.add_patch(name_box)
    ax.text(8, 68.6, "B2", fontsize=10, fontweight='bold', ha='center', va='center', color="#201F1E")
    
    # fx symbol
    ax.text(14.5, 68.6, "fx", fontsize=11, fontstyle='italic', fontweight='bold', color="#605E5C", va='center')
    
    # Formula input bar
    formula_bar = patches.Rectangle((17, 66.5), 80, 4.2, fc="#FFFFFF", ec="#8A8886")
    ax.add_patch(formula_bar)
    ax.text(19, 68.6, "=UPPER(TRIM(A2))  <-- 在此輸入資料清洗公式或直接編輯內容", fontsize=9.5, color="#004E8C", va='center')
    
    # Grid area
    grid_top = 64
    grid_bottom = 9
    col_widths = [4, 15, 18, 12, 12, 15, 12, 10]
    col_headers = ["", "A (提單號碼)", "B (清洗後提單號)", "C (起運港)", "D (目的港)", "E (貨櫃號碼)", "F (櫃型)", "G (毛重)"]
    
    # Header row
    curr_x = 2
    for j, (w, h_text) in enumerate(zip(col_widths, col_headers)):
        c_rect = patches.Rectangle((curr_x, grid_top-4), w, 4, fc="#E1DFDD" if j>0 else "#C8C6C4", ec="#8A8886", lw=0.5)
        ax.add_patch(c_rect)
        ax.text(curr_x + w/2, grid_top-2, h_text, fontsize=8.5, fontweight='bold', ha='center', va='center', color="#323130")
        curr_x += w
        
    # Grid rows (1 to 8)
    sample_rows = [
        ["1", "  emcu2049182 ", "=UPPER(TRIM(A2))", "TWKHH", "USLAX", "EMCU8273641", "40'HC", "18,450 KGS"],
        ["2", "OOCLU918273  ", "=UPPER(TRIM(A3))", "tw tpe", "DEHAM", "OOLU1928374", "40HQ", "12,400 KGS"],
        ["3", "MSCU9928172   ", "=UPPER(TRIM(A4))", "twkhh", "uslgb", "MSCU3847291", "20GP", "9,800 KGS"],
        ["4", "YMLU0192837   ", "=UPPER(TRIM(A5))", "TWTXG", "JPTYO", "YMLU5566778", "20'GP", "5,600 KGS"],
        ["5", "WHLU8837261   ", "=UPPER(TRIM(A6))", "TWKEL", "SGSIN", "WHLU9900112", "40 DRY", "21,300 KGS"],
        ["6", "CMAU7718293   ", "=UPPER(TRIM(A7))", "twkhh", "FRFOS", "CMAU4455667", "40HQ", "16,200 KGS"],
        ["7", "ONEU6627182   ", "=UPPER(TRIM(A8))", "TWTPE", "USOAK", "ONEU1122334", "40HC", "19,500 KGS"],
        ["8", "EMCU5519283   ", "=UPPER(TRIM(A9))", "TWKHH", "NLRTM", "EMCU9988776", "20GP", "8,950 KGS"],
    ]
    
    y_row = grid_top - 4
    for r_idx, row_data in enumerate(sample_rows):
        y_row -= 4.5
        curr_x = 2
        for c_idx, (w, val) in enumerate(zip(col_widths, row_data)):
            # highlight active cell B2
            is_active = (r_idx == 0 and c_idx == 2)
            bg_c = "#FFFFFF" if not is_active else "#E2F0D9"
            border_c = "#8A8886" if not is_active else "#107C41"
            lw = 0.5 if not is_active else 2.0
            
            c_rect = patches.Rectangle((curr_x, y_row), w, 4.5, fc=bg_c, ec=border_c, lw=lw)
            ax.add_patch(c_rect)
            
            # Text alignment
            align = 'center' if c_idx in [0, 3, 4, 6] else ('left' if c_idx in [1, 2, 5] else 'right')
            x_txt = curr_x + w/2 if align == 'center' else (curr_x + 1 if align == 'left' else curr_x + w - 1)
            t_color = "#107C41" if is_active else "#201F1E"
            weight = 'bold' if is_active or c_idx == 0 else 'normal'
            ax.text(x_txt, y_row + 2.2, val, fontsize=8, ha=align, va='center', color=t_color, fontweight=weight)
            curr_x += w
            
    # Sheet tab bar
    sheet_bar = patches.Rectangle((1, 4.5), 98, 4.5, fc="#F3F2F1", ec="#E1DFDD")
    ax.add_patch(sheet_bar)
    tab1 = patches.Rectangle((3, 4.5), 18, 4.5, fc="#FFFFFF", ec="#E1DFDD")
    ax.add_patch(tab1)
    ax.plot([3, 21], [4.7, 4.7], color="#107C41", lw=2)
    ax.text(12, 6.7, "01_待處理原始資料", fontsize=8.5, ha='center', va='center', color="#107C41", fontweight='bold')
    
    tab2 = patches.Rectangle((22, 4.5), 18, 4.5, fc="#EDEBE9", ec="#E1DFDD")
    ax.add_patch(tab2)
    ax.text(31, 6.7, "02_清洗完成資料表", fontsize=8.5, ha='center', va='center', color="#605E5C")
    
    ax.text(42, 6.7, "[ + ] 新增工作表", fontsize=8.5, va='center', color="#605E5C")
    
    # Status bar
    status_bar = patches.Rectangle((1, 1), 98, 3.5, fc="#107C41", ec="none")
    ax.add_patch(status_bar)
    ax.text(3, 2.7, "就緒 (Ready)   |   可編輯狀態   |   項目個數: 30   |   平均值: 14,025   |   加總: 112,200", fontsize=8, color="#FFFFFF", va='center')
    ax.text(94, 2.7, "100% [ - | + ]", fontsize=8, color="#FFFFFF", va='center', ha='right')
    
    # Visual Annotations & Callouts (Red Arrows and Text Labels)
    # 1. 功能區 Ribbon
    ax.annotate("1. 功能區 (Ribbon)\n集合所有選單與工具命令", xy=(48, 88), xytext=(48, 97),
                arrowprops=dict(facecolor='#D83B01', shrink=0.08, width=1.5, headwidth=6),
                bbox=dict(boxstyle="round,pad=0.3", fc="#FFF4CE", ec="#D83B01", lw=1),
                fontsize=8.5, fontweight='bold', ha='center', va='bottom', color="#A80000")
                
    # 2. 名稱方塊 Name Box
    ax.annotate("2. 名稱方塊 (Name Box)\n顯示當前游標所在儲存格座標", xy=(8, 70.8), xytext=(8, 77),
                arrowprops=dict(facecolor='#0078D4', shrink=0.08, width=1.5, headwidth=6),
                bbox=dict(boxstyle="round,pad=0.3", fc="#EFF6FC", ec="#0078D4", lw=1),
                fontsize=8, ha='center', va='bottom', color="#004E8C")

    # 3. 資料編輯列 Formula Bar
    ax.annotate("3. 資料編輯列 (Formula Bar)\n輸入或修改公式與儲存格真實內容", xy=(45, 70.8), xytext=(78, 64),
                arrowprops=dict(facecolor='#0078D4', shrink=0.08, width=1.5, headwidth=6),
                bbox=dict(boxstyle="round,pad=0.3", fc="#EFF6FC", ec="#0078D4", lw=1),
                fontsize=8, ha='center', va='top', color="#004E8C")

    # 4. 工作表分頁 Sheet Tabs
    ax.annotate("4. 工作表分頁 (Sheet Tabs)\n切換原始待處理與清洗後資料庫", xy=(12, 4.5), xytext=(12, -0.5),
                arrowprops=dict(facecolor='#107C41', shrink=0.08, width=1.5, headwidth=6),
                bbox=dict(boxstyle="round,pad=0.3", fc="#E2F0D9", ec="#107C41", lw=1),
                fontsize=8, ha='center', va='top', color="#107C41")
                
    # 5. 狀態列 Status Bar
    ax.annotate("5. 狀態列 (Status Bar)\n即時顯示選取範圍計數與快速統計", xy=(55, 1), xytext=(55, -4),
                arrowprops=dict(facecolor='#605E5C', shrink=0.08, width=1.5, headwidth=6),
                bbox=dict(boxstyle="round,pad=0.3", fc="#F3F2F1", ec="#605E5C", lw=1),
                fontsize=8, ha='center', va='top', color="#323130")

    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "fig1_excel_interface.png"), bbox_inches='tight')
    plt.close()
    print("Generated fig1_excel_interface.png")

# -------------------------------------------------------------
# Figure 2: 航運資料前處理核心五步循環 SOP (圖 2-1)
# -------------------------------------------------------------
def draw_fig2():
    fig, ax = plt.subplots(figsize=(10, 4.5), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 40)
    ax.axis('off')
    
    steps = [
        ("步驟 1", "結構檢視與備份", "保護原始資料\n確認標題列唯一\n檢查資料型態", "#2B579A"),
        ("步驟 2", "字串雜質清理", "TRIM去除多餘空格\nUPPER統一大小寫\n清除不可見隱藏字元", "#107C41"),
        ("步驟 3", "欄位原子化分割", "資料剖析拆解船名航次\n分離包裝件數與單位\n提取純數值與幣別", "#D83B01"),
        ("步驟 4", "型態規格化轉換", "統一西元日期YYYY-MM-DD\n文字數字轉運算數值\n空值與錯誤值標記防呆", "#80397B"),
        ("步驟 5", "去重比對與驗證", "移除重複提單號碼\nXLOOKUP港口代碼校對\n確認欄位滿足分析要求", "#008272")
    ]
    
    x = 2
    w = 17
    h = 30
    y = 5
    for i, (num, title, desc, color) in enumerate(steps):
        # Card Box
        box = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.5", fc="#FFFFFF", ec=color, lw=2)
        ax.add_patch(box)
        # Header banner
        banner = patches.FancyBboxPatch((x, y+h-7), w, 7, boxstyle="round,pad=0.2", fc=color, ec=color)
        ax.add_patch(banner)
        ax.text(x + w/2, y + h - 2.5, num, color="white", fontsize=9, fontweight='bold', ha='center', va='center')
        ax.text(x + w/2, y + h - 5.2, title, color="white", fontsize=9.5, fontweight='bold', ha='center', va='center')
        
        # Body text
        ax.text(x + w/2, y + 11, desc, color="#323130", fontsize=8.5, ha='center', va='center', linespacing=1.6)
        
        # Arrow to next
        if i < 4:
            ax.annotate("", xy=(x + w + 2.5, y + h/2), xytext=(x + w + 0.2, y + h/2),
                        arrowprops=dict(facecolor='#8A8886', edgecolor='none', width=2, headwidth=6, shrink=0.05))
            
        x += w + 2.8
        
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "fig2_data_preprocessing_flow.png"), bbox_inches='tight')
    plt.close()
    print("Generated fig2_data_preprocessing_flow.png")

# -------------------------------------------------------------
# Figure 3: 航運常見「髒資料」七大症狀對照圖 (圖 2-2)
# -------------------------------------------------------------
def draw_fig3():
    fig, ax = plt.subplots(figsize=(10.5, 5.5), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    # Title
    ax.text(50, 95, "航運實務常見『七大髒資料症狀』與標準清洗對策", fontsize=13, fontweight='bold', ha='center', color="#1B365D")
    
    symptoms = [
        ("症狀 1：多餘隱藏空格", "  EMCU20491823  ", "EMCU20491823", "TRIM(A2) 去除首尾空白，防VLOOKUP比對失敗"),
        ("症狀 2：大小寫混亂不一", "twkhh / UsLaX", "TWKHH / USLAX", "UPPER(A2) 強制轉大寫，符合國際UN/LOCODE"),
        ("症狀 3：複合欄位未原子化", "EVER GIVEN / 0124-055W", "船名: EVER GIVEN | 航次: 0124-055W", "「資料剖析」或 LEFT/MID 拆分獨立維度"),
        ("症狀 4：數值夾雜計量單位", " 18,450.50 KGS ", "18450.50 (純數字)", "SUBSTITUTE 去除單位與逗點，轉化為運算值"),
        ("症狀 5：日期格式五花八門", "2026.11.05 / 26-11-05", "2026-11-05 (ISO標準)", "資料剖析或 DATE 函數統一西元標準日期"),
        ("症狀 6：非標準代碼與縮寫", "40'HC / 40HQ / 40'DC", "40HQ (標準40呎高櫃)", "XLOOKUP 或 IF 多條件重組歸納標準櫃型"),
        ("症狀 7：重複提單與登錄", "同一提單多次出現不同列", "唯一獨立提單記錄", "「移除重複項」或 COUNTIF 篩選保留唯一筆")
    ]
    
    y = 86
    h_row = 10
    
    # Table Header
    ax.add_patch(patches.Rectangle((2, y), 96, 7, fc="#1F497D", ec="none"))
    ax.text(16, y+3.5, "異常資料特徵 (Dirty Data)", color="white", fontsize=9.5, fontweight='bold', ha='center', va='center')
    ax.text(41, y+3.5, "航運實務典型範例", color="white", fontsize=9.5, fontweight='bold', ha='center', va='center')
    ax.text(63, y+3.5, "清洗後標準結果 (Clean)", color="white", fontsize=9.5, fontweight='bold', ha='center', va='center')
    ax.text(85, y+3.5, "Excel 解決方案 / 函數", color="white", fontsize=9.5, fontweight='bold', ha='center', va='center')
    
    for i, (sym, dirty, clean, sol) in enumerate(symptoms):
        y -= 11
        bg = "#F2F5F9" if i % 2 == 0 else "#FFFFFF"
        ax.add_patch(patches.Rectangle((2, y), 96, 10, fc=bg, ec="#D9D9D9", lw=0.8))
        
        # Text
        ax.text(4, y+5, sym, color="#A80000", fontsize=9, fontweight='bold', va='center')
        ax.text(41, y+5, dirty, color="#C00000", fontsize=8.5, ha='center', va='center')
        ax.text(63, y+5, clean, color="#107C41", fontsize=8.5, ha='center', va='center', fontweight='bold')
        ax.text(74, y+5, sol, color="#201F1E", fontsize=8, va='center')
        
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "fig3_clean_vs_dirty_data.png"), bbox_inches='tight')
    plt.close()
    print("Generated fig3_clean_vs_dirty_data.png")

# -------------------------------------------------------------
# Figure 4: Excel 資料剖析精靈對話方塊示意圖 (圖 4-1)
# -------------------------------------------------------------
def draw_fig4():
    fig, ax = plt.subplots(figsize=(9.5, 6), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    # Dialog Window Box
    dlg = patches.FancyBboxPatch((5, 5), 90, 90, boxstyle="round,pad=0.3", fc="#FFFFFF", ec="#0078D4", lw=2)
    ax.add_patch(dlg)
    
    # Dialog Title
    title_bar = patches.Rectangle((5, 88), 90, 7, fc="#0078D4", ec="none")
    ax.add_patch(title_bar)
    ax.text(8, 91.5, "轉換文字給資料行精靈 - 步驟 2/3 (Text to Columns Wizard)", color="white", fontsize=10.5, fontweight='bold', va='center')
    ax.text(92, 91.5, "[ X ]", color="white", fontsize=10, va='center')
    
    # Instructions
    ax.text(8, 83, "這個對話方塊可讓您設定分隔符號 (Delimiters)。您可以在下方的預覽視窗中查看資料剖析後的狀態。", fontsize=8.5, color="#323130")
    
    # Delimiters Frame
    delim_frame = patches.Rectangle((8, 52), 40, 27, fc="#FAF9F8", ec="#8A8886", lw=1)
    ax.add_patch(delim_frame)
    ax.text(12, 76, "分隔符號 (Delimiters)", fontsize=9, fontweight='bold', color="#201F1E")
    
    # Checkboxes
    checks = ["[  ] Tab 鍵 (Tab)", "[  ] 分號 (Semicolon)", "[  ] 逗號 (Comma)", "[  ] 空格 (Space)", "[ √ ] 其他 (Other):  [  /  ]"]
    for idx, c in enumerate(checks):
        col_c = "#D83B01" if "[ √ ]" in c else "#323130"
        weight = 'bold' if "[ √ ]" in c else 'normal'
        ax.text(12, 71 - idx*4, c, fontsize=8.5, color=col_c, fontweight=weight)
        
    # Treat consecutive delimiters as one
    ax.text(8, 47, "[ √ ] 視連續分隔符號為單一符號", fontsize=8.5, color="#323130")
    ax.text(8, 43, "文字辨識符號:  [ \" ]", fontsize=8.5, color="#323130")
    
    # Preview Window
    prev_box = patches.Rectangle((8, 16), 84, 24, fc="#FFFFFF", ec="#8A8886", lw=1)
    ax.add_patch(prev_box)
    ax.text(10, 37, "資料預覽 (Data preview)：", fontsize=8.5, fontweight='bold', color="#605E5C")
    
    # Sub-columns in preview
    col1 = patches.Rectangle((10, 18), 38, 16, fc="#EFF6FC", ec="#0078D4", lw=1)
    col2 = patches.Rectangle((50, 18), 40, 16, fc="#EFF6FC", ec="#0078D4", lw=1)
    ax.add_patch(col1)
    ax.add_patch(col2)
    
    ax.text(29, 31, "欄位 1: 船名 (Vessel)", fontsize=8.5, fontweight='bold', color="#004E8C", ha='center')
    ax.text(70, 31, "欄位 2: 航次 (Voyage)", fontsize=8.5, fontweight='bold', color="#004E8C", ha='center')
    ax.plot([10, 48], [28.5, 28.5], color="#0078D4", lw=0.8)
    ax.plot([50, 90], [28.5, 28.5], color="#0078D4", lw=0.8)
    
    preview_rows = [
        ("EVER GIVEN", "0124-055W"),
        ("OOCL ROTTERDAM", "088E"),
        ("MSC ISABELLA", "2311W"),
        ("WAN HAI 315", "W082")
    ]
    for r_i, (v1, v2) in enumerate(preview_rows):
        ax.text(29, 25.5 - r_i*2.2, v1, fontsize=8, ha='center', color="#201F1E")
        ax.text(70, 25.5 - r_i*2.2, v2, fontsize=8, ha='center', color="#201F1E")
        
    # Buttons
    ax.add_patch(patches.Rectangle((42, 8), 14, 5, fc="#EDEBE9", ec="#8A8886"))
    ax.text(49, 10.5, "< 上一步", fontsize=8.5, ha='center', va='center')
    ax.add_patch(patches.Rectangle((58, 8), 16, 5, fc="#0078D4", ec="#0078D4"))
    ax.text(66, 10.5, "下一步 >", color="white", fontsize=8.5, fontweight='bold', ha='center', va='center')
    ax.add_patch(patches.Rectangle((76, 8), 16, 5, fc="#EDEBE9", ec="#8A8886"))
    ax.text(84, 10.5, "完成 (Finish)", fontsize=8.5, ha='center', va='center')

    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "fig4_text_to_columns_wizard.png"), bbox_inches='tight')
    plt.close()
    print("Generated fig4_text_to_columns_wizard.png")

# -------------------------------------------------------------
# Figure 5: Excel 移除重複項對話方塊 (圖 4-2)
# -------------------------------------------------------------
def draw_fig5():
    fig, ax = plt.subplots(figsize=(8.5, 5.5), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    # Dialog Box
    dlg = patches.FancyBboxPatch((8, 8), 84, 84, boxstyle="round,pad=0.3", fc="#FFFFFF", ec="#D83B01", lw=2)
    ax.add_patch(dlg)
    
    # Title Bar
    title_bar = patches.Rectangle((8, 84), 84, 8, fc="#D83B01", ec="none")
    ax.add_patch(title_bar)
    ax.text(12, 88, "移除重複項 (Remove Duplicates)", color="white", fontsize=11, fontweight='bold', va='center')
    ax.text(88, 88, "[ X ]", color="white", fontsize=10, va='center')
    
    # Description
    ax.text(12, 79, "若要刪除重複的值，請選取一或多個包含重複項目的資料欄：", fontsize=8.5, color="#323130")
    
    # Header check
    ax.text(12, 73, "[ √ ] 我的資料有標題 (My data has headers)", fontsize=8.5, fontweight='bold', color="#201F1E")
    
    # Select all / unselect all buttons
    ax.add_patch(patches.Rectangle((12, 65), 18, 5, fc="#EDEBE9", ec="#8A8886"))
    ax.text(21, 67.5, "全選 (Select All)", fontsize=8, ha='center', va='center')
    ax.add_patch(patches.Rectangle((32, 65), 20, 5, fc="#EDEBE9", ec="#8A8886"))
    ax.text(42, 67.5, "取消全選 (Unselect)", fontsize=8, ha='center', va='center')
    
    # Column List Box
    list_box = patches.Rectangle((12, 22), 76, 40, fc="#FAF9F8", ec="#8A8886")
    ax.add_patch(list_box)
    
    cols = [
        ("[ √ ]  提單號碼_BL_NO  (★唯一主鍵比對關鍵)", True),
        ("[    ]  船名航次_VESSEL_VOY", False),
        ("[    ]  起運港_POL", False),
        ("[    ]  目的港_POD", False),
        ("[    ]  貨櫃號碼_CNTR_NO", False),
        ("[    ]  包裝件數_PKGS", False),
    ]
    for idx, (c_name, is_key) in enumerate(cols):
        y_c = 57 - idx*6
        c_color = "#A80000" if is_key else "#323130"
        weight = 'bold' if is_key else 'normal'
        ax.text(16, y_c, c_name, fontsize=8.5, color=c_color, fontweight=weight)
        
    # Buttons OK / Cancel
    ax.add_patch(patches.Rectangle((54, 12), 16, 6, fc="#0078D4", ec="#0078D4"))
    ax.text(62, 15, "確定 (OK)", color="white", fontsize=9, fontweight='bold', ha='center', va='center')
    ax.add_patch(patches.Rectangle((72, 12), 16, 6, fc="#EDEBE9", ec="#8A8886"))
    ax.text(80, 15, "取消 (Cancel)", fontsize=9, ha='center', va='center')

    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "fig5_remove_duplicates_dialog.png"), bbox_inches='tight')
    plt.close()
    print("Generated fig5_remove_duplicates_dialog.png")

# -------------------------------------------------------------
# Figure 6: 快速填入 (Flash Fill) 與資料剖析操作對照 (圖 4-3)
# -------------------------------------------------------------
def draw_fig6():
    fig, ax = plt.subplots(figsize=(9.5, 5), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    ax.text(50, 93, "Excel「快速填入」(Flash Fill / 快速鍵 Ctrl+E) 智慧型字串拆解示範", fontsize=11.5, fontweight='bold', ha='center', color="#1B365D")
    
    # Excel Grid Mockup
    ax.add_patch(patches.Rectangle((5, 10), 90, 75, fc="#FFFFFF", ec="#107C41", lw=1.5))
    
    # Headers
    h_y = 75
    ax.add_patch(patches.Rectangle((5, h_y), 90, 10, fc="#E1DFDD", ec="#8A8886"))
    ax.text(12, h_y+5, "列號", fontsize=9, fontweight='bold', ha='center', va='center')
    ax.text(32, h_y+5, "欄 A: 原始船名航次 (複合欄位)", fontsize=9, fontweight='bold', ha='center', va='center')
    ax.text(62, h_y+5, "欄 B: 提取船名 (手打範例後按Ctrl+E)", fontsize=9, fontweight='bold', ha='center', va='center', color="#004E8C")
    ax.text(85, h_y+5, "欄 C: 提取航次", fontsize=9, fontweight='bold', ha='center', va='center', color="#004E8C")
    
    grid_data = [
        ("1", "EVER GIVEN / 0124-055W", "EVER GIVEN (手動輸入基準)", "0124-055W", True),
        ("2", "OOCL ROTTERDAM / 088E", "OOCL ROTTERDAM (自動生成)", "088E", False),
        ("3", "MSC ISABELLA / 2311W", "MSC ISABELLA (自動生成)", "2311W", False),
        ("4", "YM WELLHEAD / 221E", "YM WELLHEAD (自動生成)", "221E", False),
        ("5", "WAN HAI 315 / W082", "WAN HAI 315 (自動生成)", "W082", False)
    ]
    
    y = 75
    for r, orig, b_val, c_val, is_seed in grid_data:
        y -= 12
        bg = "#FFF4CE" if is_seed else "#EBF3FC"
        ax.add_patch(patches.Rectangle((5, y), 90, 12, fc=bg, ec="#D9D9D9", lw=0.8))
        ax.text(12, y+6, r, fontsize=8.5, ha='center', va='center', color="#605E5C")
        ax.text(32, y+6, orig, fontsize=8.5, ha='center', va='center')
        ax.text(62, y+6, b_val, fontsize=8.5, ha='center', va='center', color="#107C41" if not is_seed else "#A80000", fontweight='bold')
        ax.text(85, y+6, c_val, fontsize=8.5, ha='center', va='center', color="#004E8C")
        
    ax.text(50, 4, "★ 說明：只需在第一列手動鍵入想要的結果，游標移到下一列按下 Ctrl+E，Excel將自動推斷邏輯並秒速完成全欄填充！", fontsize=8.5, ha='center', color="#D83B01", fontweight='bold')
    
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "fig6_flash_fill_demo.png"), bbox_inches='tight')
    plt.close()
    print("Generated fig6_flash_fill_demo.png")

# -------------------------------------------------------------
# Figure 7: Excel 資料驗證 (Data Validation) 防呆選單設定 (圖 4-4)
# -------------------------------------------------------------
def draw_fig7():
    fig, ax = plt.subplots(figsize=(8.5, 5.5), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    # Dialog Box
    dlg = patches.FancyBboxPatch((8, 8), 84, 84, boxstyle="round,pad=0.3", fc="#FFFFFF", ec="#107C41", lw=2)
    ax.add_patch(dlg)
    
    # Title Bar
    title_bar = patches.Rectangle((8, 84), 84, 8, fc="#107C41", ec="none")
    ax.add_patch(title_bar)
    ax.text(12, 88, "資料驗證 (Data Validation) - 防呆設定", color="white", fontsize=11, fontweight='bold', va='center')
    ax.text(88, 88, "[ X ]", color="white", fontsize=10, va='center')
    
    # Tabs
    ax.add_patch(patches.Rectangle((12, 76), 22, 6, fc="#FFFFFF", ec="#8A8886"))
    ax.text(23, 79, "設定 (Settings)", fontsize=8.5, fontweight='bold', color="#107C41", ha='center', va='center')
    ax.add_patch(patches.Rectangle((35, 76), 22, 6, fc="#EDEBE9", ec="#8A8886"))
    ax.text(46, 79, "輸入訊息 (Input)", fontsize=8.5, color="#605E5C", ha='center', va='center')
    ax.add_patch(patches.Rectangle((58, 76), 22, 6, fc="#EDEBE9", ec="#8A8886"))
    ax.text(69, 79, "錯誤提醒 (Error)", fontsize=8.5, color="#605E5C", ha='center', va='center')
    
    # Validation criteria frame
    ax.text(14, 69, "儲存格內允許的驗證準則：", fontsize=9, fontweight='bold', color="#201F1E")
    
    ax.text(14, 61, "允許 (Allow):", fontsize=8.5, color="#323130")
    ax.add_patch(patches.Rectangle((32, 58), 45, 6, fc="#FAF9F8", ec="#8A8886"))
    ax.text(35, 61, "清單 (List)  [ v ]", fontsize=8.5, fontweight='bold', color="#0078D4", va='center')
    
    ax.text(14, 49, "[ √ ] 忽略空白 (Ignore blank)", fontsize=8.5, color="#323130")
    ax.text(50, 49, "[ √ ] 提供儲存格內下拉式選單", fontsize=8.5, color="#107C41", fontweight='bold')
    
    ax.text(14, 39, "來源 (Source):", fontsize=8.5, fontweight='bold', color="#323130")
    ax.add_patch(patches.Rectangle((32, 36), 55, 6, fc="#FAF9F8", ec="#8A8886"))
    ax.text(34, 39, "20GP, 40GP, 40HQ, 45HQ, 20RF, 40RF", fontsize=8.5, color="#A80000", fontweight='bold', va='center', family='Consolas')
    
    ax.text(14, 28, "★ 航運防呆實務應用說明：\n強制鎖定櫃型欄位僅能從標準海運ISO清單中點選，\n徹底根絕行政同仁輸入 40'HC、40DC 等非標準縮寫！", fontsize=8, color="#004E8C", linespacing=1.4)
    
    # Buttons
    ax.add_patch(patches.Rectangle((54, 12), 16, 6, fc="#107C41", ec="#107C41"))
    ax.text(62, 15, "確定 (OK)", color="white", fontsize=9, fontweight='bold', ha='center', va='center')
    ax.add_patch(patches.Rectangle((72, 12), 16, 6, fc="#EDEBE9", ec="#8A8886"))
    ax.text(80, 15, "取消 (Cancel)", fontsize=9, ha='center', va='center')

    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "fig7_data_validation_dialog.png"), bbox_inches='tight')
    plt.close()
    print("Generated fig7_data_validation_dialog.png")

# -------------------------------------------------------------
# Figure 8: 前處理十大核心函數架構圖 (圖 4-5)
# -------------------------------------------------------------
def draw_fig8():
    fig, ax = plt.subplots(figsize=(10, 5.5), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    ax.text(50, 95, "航運資料前處理『十大核心函數』體系導圖", fontsize=12.5, fontweight='bold', ha='center', color="#1B365D")
    
    categories = [
        ("字串清洗類 (Text Cleaning)", "#107C41", [
            ("TRIM", "去除字串首尾及多餘內部空格"),
            ("CLEAN", "清除系統匯出的不可列印ASCII換行字元"),
            ("UPPER / LOWER", "英文大小寫強制標準化")
        ]),
        ("文字拆解替換 (Extraction & Replace)", "#0078D4", [
            ("LEFT / MID / RIGHT", "依指定字元數精準截取航線與櫃號"),
            ("LEN / FIND", "計算字串長度並動態定位分隔符號"),
            ("SUBSTITUTE", "精確替換多餘文字、破折號與單位")
        ]),
        ("型態數值轉換 (Type Casting)", "#D83B01", [
            ("VALUE", "將帶引號之文字型數字強制轉為真數值"),
            ("TEXT", "將純數值或日期格式化為標準字串顯示")
        ]),
        ("日期與查找檢驗 (Date & Lookup)", "#80397B", [
            ("DATE / YEAR / MONTH", "重組分散的年、月、日為標準日期"),
            ("IFERROR / XLOOKUP", "動態比對港口主檔並攔截屏蔽錯誤值")
        ])
    ]
    
    x = 3
    w = 22
    for cat_title, cat_color, func_list in categories:
        # Category Box
        ax.add_patch(patches.FancyBboxPatch((x, 10), w, 78, boxstyle="round,pad=0.4", fc="#FFFFFF", ec=cat_color, lw=1.8))
        # Banner
        ax.add_patch(patches.FancyBboxPatch((x, 78), w, 10, boxstyle="round,pad=0.2", fc=cat_color, ec=cat_color))
        ax.text(x + w/2, 83, cat_title, color="white", fontsize=8.5, fontweight='bold', ha='center', va='center')
        
        y_f = 68
        for f_name, f_desc in func_list:
            f_box = patches.FancyBboxPatch((x+1, y_f-12), w-2, 16, boxstyle="round,pad=0.2", fc="#F8F9FA", ec="#D9D9D9")
            ax.add_patch(f_box)
            ax.text(x + 2, y_f+1, f_name, fontsize=9.5, fontweight='bold', color=cat_color, family='Consolas')
            ax.text(x + 2, y_f-6, f_desc, fontsize=7.5, color="#323130", linespacing=1.3)
            y_f -= 21
            
        x += w + 2.5
        
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "fig8_core_functions_architecture.png"), bbox_inches='tight')
    plt.close()
    print("Generated fig8_core_functions_architecture.png")

# -------------------------------------------------------------
# Figure 9: 30筆航運示範案例清洗前後成效對照圖 (圖 3-1)
# -------------------------------------------------------------
def draw_fig9():
    fig, ax = plt.subplots(figsize=(10.5, 6), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    ax.text(50, 96, "30筆航運示範案例：資料清洗前 (Dirty) 與 清洗後 (Clean) 成效對照", fontsize=12, fontweight='bold', ha='center', color="#1B365D")
    
    # Left Box: Dirty Raw Data
    ax.add_patch(patches.FancyBboxPatch((2, 6), 46, 85, boxstyle="round,pad=0.3", fc="#FFF0F0", ec="#C00000", lw=1.5))
    ax.add_patch(patches.Rectangle((2, 83), 46, 8, fc="#C00000", ec="none"))
    ax.text(25, 87, "[X] 清洗前：混亂髒資料 (Raw Messy)", color="white", fontsize=9.5, fontweight='bold', ha='center', va='center')
    
    dirty_items = [
        "● 提單號碼：首尾含大量空白 (如'  EMCU20491823 ')",
        "● 船名航次：合併同一欄 (如'EVER GIVEN / 0124-055W')",
        "● 起訖港口：大小寫不一致、空白混雜 ('tw tpe', ' uslgB')",
        "● 貨櫃號碼：夾雜中線連字號 ('EMCU-8273641')",
        "● 貨櫃規格：稱呼混亂 (40'HC, 40HQ, 40 DRY, 20'GP)",
        "● 件數重量：數字與英文單位混雜 ('1,200 CTNS', '18,450 KGS')",
        "● 結關開航：日期格式混雜 (2026/11/02, 2026-11-03, 2026.11.04)",
        "● 運費金額：含貨幣符號文字 ('USD 2,450.00', '$1,250.00')",
        "● 重複登錄：項次4與項次28整筆提單完全重複"
    ]
    y_d = 77
    for item in dirty_items:
        ax.text(4, y_d, item, fontsize=7.5, color="#800000", va='center')
        y_d -= 7.5
        
    # Arrow in middle
    ax.annotate("", xy=(52, 48), xytext=(48, 48),
                arrowprops=dict(facecolor='#107C41', edgecolor='none', width=3, headwidth=8))
                
    # Right Box: Clean Data
    ax.add_patch(patches.FancyBboxPatch((52, 6), 46, 85, boxstyle="round,pad=0.3", fc="#F0FFF0", ec="#107C41", lw=1.5))
    ax.add_patch(patches.Rectangle((52, 83), 46, 8, fc="#107C41", ec="none"))
    ax.text(75, 87, "[√] 清洗後：標準分析就緒 (Tidy Data)", color="white", fontsize=9.5, fontweight='bold', ha='center', va='center')
    
    clean_items = [
        "[OK] 提單號碼：TRIM+UPPER 純淨大寫無空格 (EMCU20491823)",
        "[OK] 船名航次：資料剖析原子化獨立兩欄 ('EVER GIVEN', '0124-055W')",
        "[OK] 起訖港口：標準5碼 UN/LOCODE ('TWTPE', 'USLGB')",
        "[OK] 貨櫃號碼：SUBSTITUTE 去除連字號標準11碼",
        "[OK] 貨櫃規格：標準化歸納為 20GP, 40GP, 40HQ",
        "[OK] 件數重量：拆解為純數值與獨立計量單位 (可加總運算)",
        "[OK] 結關開航：標準西元日期 YYYY-MM-DD (可計算天數航期)",
        "[OK] 運費金額：純浮點數值 2450.00 (可套用樞紐分析表統計)",
        "[OK] 唯一記錄：成功移除項次28重複項，保留29筆乾淨主檔"
    ]
    y_c = 77
    for item in clean_items:
        ax.text(54, y_c, item, fontsize=7.5, color="#006000", va='center')
        y_c -= 7.5

    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "fig9_demo30_before_after.png"), bbox_inches='tight')
    plt.close()
    print("Generated fig9_demo30_before_after.png")

# -------------------------------------------------------------
# Figure 10: 300筆實務情境題庫架構與業務流程圖 (圖 5-1)
# -------------------------------------------------------------
def draw_fig10():
    fig, ax = plt.subplots(figsize=(10, 4.8), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    ax.text(50, 93, "300筆航運情境練習題庫：四大業務情境模組架構圖", fontsize=12, fontweight='bold', ha='center', color="#1B365D")
    
    modules = [
        ("模組 A (100筆)", "海運出口訂艙裝船清單", "提單號清洗、貨櫃ISO規格化\n件數重量單位拆分、日期規範", "#1F497D"),
        ("模組 B (70筆)", "貨櫃滯期延滯費 (D&D)", "卸船還櫃天數運算、空值判斷\n逾期天數MAX截斷、計費金額", "#D83B01"),
        ("模組 C (70筆)", "貨代應收應付帳單", "外幣幣別分離、數值型態轉化\n財務匯率換算、5%稅額精算", "#274E13"),
        ("模組 D (60筆)", "跨國港口與船期標準化", "UN/LOCODE國際碼動態查找\n別名轉換、船期航程天數計算", "#6A1B9A")
    ]
    
    x = 3
    w = 22
    for code, title, details, col in modules:
        ax.add_patch(patches.FancyBboxPatch((x, 15), w, 70, boxstyle="round,pad=0.3", fc="#FFFFFF", ec=col, lw=1.8))
        ax.add_patch(patches.FancyBboxPatch((x, 70), w, 15, boxstyle="round,pad=0.2", fc=col, ec=col))
        ax.text(x + w/2, 79, code, color="white", fontsize=9, fontweight='bold', ha='center', va='center')
        ax.text(x + w/2, 73, title, color="white", fontsize=9, fontweight='bold', ha='center', va='center')
        
        ax.text(x + w/2, 42, details, color="#323130", fontsize=8, ha='center', va='center', linespacing=1.6)
        
        # Bottom tag
        ax.add_patch(patches.Rectangle((x+2, 18), w-4, 6, fc="#F3F2F1", ec="#8A8886"))
        ax.text(x + w/2, 21, "附完整解題指引", fontsize=7.5, color="#605E5C", ha='center', va='center', fontweight='bold')
        
        x += w + 2.5
        
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "fig10_scenario_300_overview.png"), bbox_inches='tight')
    plt.close()
    print("Generated fig10_scenario_300_overview.png")

print("Rendering all diagrams...")
draw_fig1()
draw_fig2()
draw_fig3()
draw_fig4()
draw_fig5()
draw_fig6()
draw_fig7()
draw_fig8()
draw_fig9()
draw_fig10()
print("All 10 figures rendered successfully!")
