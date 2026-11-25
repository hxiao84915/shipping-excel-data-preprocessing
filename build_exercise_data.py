import os
import random
import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

random.seed(101)

def generate_300_records():
    # Module A: Records 1 to 100 (出口訂艙裝船清單)
    mod_a_rows = []
    carriers = ["EMC", "OOCL", "MSC", "YML", "WHL", "CMA", "ONE", "HLC", "SITC"]
    ports_pol = ["TWKHH", "TWTPE", "TWKEL", "TWTXG"]
    ports_pod = ["USLAX", "DEHAM", "NLRTM", "SGSIN", "JPTYO", "GBFLX", "AUMEL", "USOAK", "VNSGN", "CNSHA"]
    shippers = [
        "GREAT OCEAN TRADING CO.", "APEX LOGISTICS CORP", "FORMOSA PLASTICS CORP", 
        "CHENG SHIN RUBBER IND.", "KING CAR FOOD IND. CO.", "UNI-PRESIDENT ENTERPRISES", 
        "GIANT MANUFACTURING CO.", "DELTA ELECTRONICS INC.", "TATUNG COMPANY LTD", 
        "TAIWAN CEMENT CORP.", "ACER INCORPORATED", "ASUSTEK COMPUTER INC.", 
        "NAN YA PLASTICS CORP.", "QUANTA COMPUTER INC.", "COMPAL ELECTRONICS INC", 
        "AU OPTRONICS CORP.", "INVENTEC CORPORATION", "WISTRON CORPORATION", 
        "PEGATRON CORPORATION", "MEDIATEK INC.", "FOXCONN TECHNOLOGY GROUP", 
        "FAR EASTERN NEW CENTURY", "EVERGREEN STEEL CORP.", "TSMC LOGISTICS DIV"
    ]
    types_raw = ["40'HC", "40HQ", "20GP", "20'GP", "40 DRY", "40FR", "20 OT", "40'DC"]
    
    start_date = datetime.date(2026, 10, 1)
    
    # 100 records for Module A
    for i in range(1, 101):
        bl_prefix = random.choice(carriers)
        bl_num = f"{bl_prefix}{random.randint(10000000, 99999999)}"
        # Inject dirtiness into BL:
        if i % 7 == 0:
            bl_str = f"  {bl_num} "
        elif i % 11 == 0:
            bl_str = bl_num.lower()
        else:
            bl_str = bl_num
            
        vessel = f"VESSEL {chr(65 + (i % 26))} / {random.randint(100, 999)}{random.choice(['E', 'W', 'N', 'S'])}"
        
        pol = random.choice(ports_pol)
        if i % 5 == 0:
            pol_str = pol.lower()
        elif i % 8 == 0:
            pol_str = f" {pol} "
        else:
            pol_str = pol
            
        pod = random.choice(ports_pod)
        if i % 6 == 0:
            pod_str = pod.lower()
        elif i % 9 == 0:
            pod_str = f"  {pod}"
        else:
            pod_str = pod
            
        cntr = f"{random.choice(['EMCU', 'OOLU', 'MSCU', 'YMLU', 'WHLU', 'TGHU'])}{random.randint(1000000, 9999999)}"
        if i % 4 == 0:
            cntr_str = f"{cntr[:4]}-{cntr[4:]}"
        elif i % 7 == 0:
            cntr_str = f" {cntr.lower()} "
        else:
            cntr_str = cntr
            
        ctype = random.choice(types_raw)
        
        # Pkgs
        pkgs_num = random.randint(100, 2500)
        p_unit = random.choice(["CTNS", "PKGS", "PLTS", "BAGS"])
        if i % 3 == 0:
            pkgs_str = f"{pkgs_num:,} {p_unit}"
        elif i % 5 == 0:
            pkgs_str = f"{pkgs_num}"
        else:
            pkgs_str = f" {pkgs_num} {p_unit} "
            
        # Weight
        wt = round(random.uniform(4000, 26000), 2)
        if i % 2 == 0:
            wt_str = f" {wt:,.2f} KGS "
        elif i % 3 == 0:
            wt_str = f"{wt:.1f}"
        else:
            wt_str = f"{wt:,.0f}KGS"
            
        # CBM
        cbm = round(random.uniform(15.0, 68.0), 2)
        if i % 2 == 0:
            cbm_str = f"{cbm:.2f} CBM"
        elif i % 4 == 0:
            cbm_str = f" {cbm:.1f} "
        else:
            cbm_str = f"{cbm:.2f}"
            
        # Date
        d_cutoff = start_date + datetime.timedelta(days=(i * 2) % 60)
        d_etd = d_cutoff + datetime.timedelta(days=random.randint(3, 5))
        
        # Format date with mixed formats
        def dirty_date(d, seed):
            if seed % 4 == 0:
                return d.strftime("%Y/%m/%d")
            elif seed % 4 == 1:
                return d.strftime("%Y-%m-%d")
            elif seed % 4 == 2:
                return d.strftime("%Y.%m.%d")
            else:
                return d.strftime("%y/%m/%d")
                
        cut_str = dirty_date(d_cutoff, i)
        etd_str = dirty_date(d_etd, i + 1)
        
        terms = random.choice(["PREPAID", "COLLECT", "prepaid", "collect", " Prepaid ", " Collect "])
        
        amt = random.randint(800, 4500)
        if i % 3 == 0:
            amt_str = f"USD {amt:,.2f}"
        elif i % 4 == 0:
            amt_str = f"${amt:,}"
        elif i % 5 == 0:
            amt_str = f" {amt}.00 "
        else:
            amt_str = f"{amt:,}"
            
        shipper = random.choice(shippers)
        if i % 3 == 0:
            shipper_str = f"  {shipper}  "
        elif i % 6 == 0:
            shipper_str = shipper.lower()
        else:
            shipper_str = shipper
            
        mod_a_rows.append({
            "模組分類": "A.訂艙與裝船清單清理",
            "項次": i,
            "業務編號": bl_str,
            "關聯識別碼": cntr_str,
            "規格項目": ctype,
            "主體名稱": vessel,
            "起點_POL": pol_str,
            "迄點_POD": pod_str,
            "量值1_件數": pkgs_str,
            "量值2_重量": wt_str,
            "量值3_材積": cbm_str,
            "起始基準日": cut_str,
            "結束航行日": etd_str,
            "分類條款": terms,
            "費用金額": amt_str,
            "關係客戶": shipper_str
        })
        
    # Module B: Records 101 to 170 (貨櫃滯期延滯費 D&D 逾期計算前處理)
    mod_b_rows = []
    for i in range(101, 171):
        cntr = f"{random.choice(['EMCU', 'OOLU', 'MSCU', 'YMLU', 'WHLU', 'TGHU'])}{random.randint(1000000, 9999999)}"
        ctype = random.choice(["20GP", "40GP", "40HQ", "20'GP", "40'HC"])
        vessel = f"PACIFIC VOYAGER / {random.randint(200, 300)}W"
        pol = "TWKHH"
        pod = random.choice(["USLAX", "DEHAM", "JPTYO", "SGSIN"])
        
        d_disch = start_date + datetime.timedelta(days=(i - 100) * 2)
        used_days = random.randint(3, 25)
        d_return = d_disch + datetime.timedelta(days=used_days)
        free_days = random.choice([7, 10, 14])
        
        daily_rate = 50.0 if "20" in ctype else 90.0
        
        def dirty_d(d, s):
            if s % 3 == 0: return d.strftime("%Y/%m/%d")
            elif s % 3 == 1: return d.strftime("%Y-%m-%d")
            else: return d.strftime("%Y.%m.%d")
            
        disch_str = dirty_d(d_disch, i)
        ret_str = dirty_d(d_return, i + 2) if (i % 15 != 0) else "" # some unreturned
        
        mod_b_rows.append({
            "模組分類": "B.貨櫃延滯費(D&D)分析",
            "項次": i,
            "業務編號": f"DD-INV-{2026000 + i}",
            "關聯識別碼": cntr,
            "規格項目": ctype,
            "主體名稱": f"免費天數: {free_days}天",
            "起點_POL": pol,
            "迄點_POD": pod,
            "量值1_件數": f"費率: USD {daily_rate}/天",
            "量值2_重量": f"約定免收: {free_days}天",
            "量值3_材積": f"日租率 {daily_rate}",
            "起始基準日": disch_str, # 卸船/提櫃日
            "結束航行日": ret_str,   # 還空櫃日
            "分類條款": "IMPORT D&D" if i % 2 == 0 else "EXPORT DETENTION",
            "費用金額": f"USD {daily_rate}",
            "關係客戶": random.choice(shippers)
        })

    # Module C: Records 171 to 240 (貨代帳單幣別匯率金額清理)
    mod_c_rows = []
    fee_items = ["海運基本運費(O/F)", "港口碼頭作業費(THC)", "提單製作費(DOC)", "貨櫃封條費(SEAL)", "進出口報關費(CUST)", "內陸卡車拖運費(TRK)"]
    currencies = [("USD", 32.5), ("TWD", 1.0), ("EUR", 35.2), ("JPY", 0.215), ("RMB", 4.5)]
    
    for i in range(171, 241):
        inv_no = f"INV2026-{i:04d}"
        bl_ref = f"BL-TW-{random.randint(10000, 99999)}"
        fee = random.choice(fee_items)
        curr, rate = random.choice(currencies)
        
        if curr == "USD":
            raw_amt = random.randint(150, 3200)
            amt_display = f"USD {raw_amt:,.2f}" if i % 2 == 0 else f"${raw_amt}"
        elif curr == "TWD":
            raw_amt = random.randint(1500, 45000)
            amt_display = f"NT$ {raw_amt:,}" if i % 2 == 0 else f"{raw_amt:,} TWD"
        elif curr == "EUR":
            raw_amt = random.randint(200, 2800)
            amt_display = f"EUR {raw_amt:,.2f}"
        elif curr == "JPY":
            raw_amt = random.randint(10000, 150000)
            amt_display = f"JPY {raw_amt:,}"
        else: # RMB
            raw_amt = random.randint(500, 8000)
            amt_display = f"RMB {raw_amt:,.2f}"
            
        mod_c_rows.append({
            "模組分類": "C.貨代應收應付帳單清理",
            "項次": i,
            "業務編號": inv_no,
            "關聯識別碼": bl_ref,
            "規格項目": fee,
            "主體名稱": f"匯率基準: {rate}",
            "起點_POL": "TWTPE",
            "迄點_POD": "GLOBAL",
            "量值1_件數": curr,
            "量值2_重量": str(rate),
            "量值3_材積": str(raw_amt),
            "起始基準日": (start_date + datetime.timedelta(days=(i - 170))).strftime("%Y/%m/%d"),
            "結束航行日": (start_date + datetime.timedelta(days=(i - 170) + 30)).strftime("%Y/%m/%d"), # 付款截止
            "分類條款": "AR-應收" if i % 2 == 0 else "AP-應付",
            "費用金額": f" {amt_display} ",
            "關係客戶": random.choice(shippers)
        })

    # Module D: Records 241 to 300 (跨國港口代碼與船期表標準化)
    mod_d_rows = []
    port_dict = [
        ("Kaohsiung", "TWKHH", "台灣高雄港", "Asia"),
        ("Keelung", "TWKEL", "台灣基隆港", "Asia"),
        ("Taichung", "TWTXG", "台灣台中港", "Asia"),
        ("Taipei", "TWTPE", "台灣台北港", "Asia"),
        ("Los Angeles", "USLAX", "美國洛杉磯港", "North America"),
        ("Long Beach", "USLGB", "美國長灘港", "North America"),
        ("Oakland", "USOAK", "美國奧克蘭港", "North America"),
        ("Hamburg", "DEHAM", "德國漢堡港", "Europe"),
        ("Rotterdam", "NLRTM", "荷蘭鹿特丹港", "Europe"),
        ("Felixstowe", "GBFLX", "英國費利克斯托港", "Europe"),
        ("Singapore", "SGSIN", "新加坡港", "Southeast Asia"),
        ("Tokyo", "JPTYO", "日本東京港", "Asia"),
        ("Shanghai", "CNSHA", "中國上海港", "Asia"),
        ("Piraeus", "GRPIR", "希臘比雷埃夫斯港", "Europe")
    ]
    
    for i in range(241, 301):
        p_name, un_code, p_cname, region = random.choice(port_dict)
        service = random.choice(["PSW-美國西南線", "AEU-歐地直航線", "KIX-關西快捷線", "AGX-中東航線", "AUE-澳洲東岸線"])
        vsl = f"GLOBAL EXPRESS / {random.randint(100, 150)}{random.choice(['E', 'W'])}"
        
        # Inject messy port code or name
        if i % 3 == 0:
            messy_code = un_code.lower()
        elif i % 4 == 0:
            messy_code = f"  {un_code} "
        elif i % 5 == 0:
            messy_code = un_code.replace("TW", "T-")
        else:
            messy_code = un_code
            
        eta_d = start_date + datetime.timedelta(days=(i - 240) * 3)
        etd_d = eta_d + datetime.timedelta(days=random.randint(1, 3))
        
        mod_d_rows.append({
            "模組分類": "D.跨國港口與船期標準化",
            "項次": i,
            "業務編號": f"SCH-2026-{i:03d}",
            "關聯識別碼": messy_code,
            "規格項目": service,
            "主體名稱": vsl,
            "起點_POL": "TWKHH",
            "迄點_POD": p_name,
            "量值1_件數": region,
            "量值2_重量": p_cname,
            "量值3_材積": un_code,
            "起始基準日": eta_d.strftime("%Y/%m/%d") if i % 2 == 0 else eta_d.strftime("%Y-%m-%d"),
            "結束航行日": etd_d.strftime("%Y/%m/%d") if i % 2 == 0 else etd_d.strftime("%Y-%m-%d"),
            "分類條款": "DIRECT-直航" if i % 3 != 0 else "TS-轉運",
            "費用金額": f"標準費率: USD {random.randint(1200, 3500)}",
            "關係客戶": "CARRIER ALLIANCE CO."
        })
        
    all_300 = mod_a_rows + mod_b_rows + mod_c_rows + mod_d_rows
    df_300 = pd.DataFrame(all_300)
    return df_300

print("Generating 300 comprehensive shipping exercise records...")
df_exercise_300 = generate_300_records()
print(f"Exercise 300 records shape: {df_exercise_300.shape}")

# Save to Excel
wb_ex = openpyxl.Workbook()
ws_ex_raw = wb_ex.active
ws_ex_raw.title = "01_300筆航運情境練習題庫(Raw)"

# Add lookup tables sheet
ws_lookup = wb_ex.create_sheet(title="02_基礎代碼對照表(Lookup)")
ws_keys = wb_ex.create_sheet(title="03_解題思路與公式指引(Solutions)")

# Setup Styles
header_fill_dark = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
header_font_white = Font(name="Microsoft JhengHei", size=11, bold=True, color="FFFFFF")
cell_font = Font(name="Microsoft JhengHei", size=10)
border_grid = Border(
    left=Side(style='thin', color='E0E0E0'),
    right=Side(style='thin', color='E0E0E0'),
    top=Side(style='thin', color='E0E0E0'),
    bottom=Side(style='thin', color='E0E0E0')
)
zebra_stripe = PatternFill(start_color="F7F9FB", end_color="F7F9FB", fill_type="solid")

# Write Raw Sheet
headers = list(df_exercise_300.columns)
ws_ex_raw.append(headers)
for col_idx in range(1, len(headers) + 1):
    c = ws_ex_raw.cell(row=1, column=col_idx)
    c.fill = header_fill_dark
    c.font = header_font_white
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

for r_idx, row in enumerate(df_exercise_300.values, start=2):
    ws_ex_raw.append(list(row))
    for col_idx in range(1, len(headers) + 1):
        c = ws_ex_raw.cell(row=r_idx, column=col_idx)
        c.font = cell_font
        c.border = border_grid
        if r_idx % 2 == 1:
            c.fill = zebra_stripe
        if col_idx in [1, 2, 7, 8, 12, 13, 14]:
            c.alignment = Alignment(horizontal="center")

for col in ws_ex_raw.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws_ex_raw.column_dimensions[col_letter].width = min(max(max_len * 1.3, 12), 35)

# Write Lookup Table
lookup_headers = ["港口英文名稱", "UN/LOCODE國際代碼", "中文港口名稱", "所屬航線區域", "預設免費天數", "滯期費率(USD/天)"]
lookup_data = [
    ["Kaohsiung", "TWKHH", "台灣高雄港", "遠東/亞洲", 14, 50],
    ["Keelung", "TWKEL", "台灣基隆港", "遠東/亞洲", 14, 50],
    ["Taichung", "TWTXG", "台灣台中港", "遠東/亞洲", 14, 50],
    ["Taipei", "TWTPE", "台灣台北港", "遠東/亞洲", 14, 50],
    ["Los Angeles", "USLAX", "美國洛杉磯港", "美西航線", 7, 90],
    ["Long Beach", "USLGB", "美國長灘港", "美西航線", 7, 90],
    ["Oakland", "USOAK", "美國奧克蘭港", "美西航線", 7, 90],
    ["Hamburg", "DEHAM", "德國漢堡港", "歐洲航線", 10, 85],
    ["Rotterdam", "NLRTM", "荷蘭鹿特丹港", "歐洲航線", 10, 85],
    ["Felixstowe", "GBFLX", "英國費利克斯托港", "歐洲航線", 10, 85],
    ["Singapore", "SGSIN", "新加坡港", "東南亞航線", 10, 60],
    ["Tokyo", "JPTYO", "日本東京港", "亞洲/日本線", 14, 55],
    ["Shanghai", "CNSHA", "中國上海港", "兩岸/亞洲", 14, 50],
    ["Piraeus", "GRPIR", "希臘比雷埃夫斯港", "地中海航線", 10, 75]
]
ws_lookup.append(lookup_headers)
for col_idx in range(1, len(lookup_headers) + 1):
    c = ws_lookup.cell(row=1, column=col_idx)
    c.fill = PatternFill(start_color="274E13", end_color="274E13", fill_type="solid")
    c.font = header_font_white
    c.alignment = Alignment(horizontal="center", vertical="center")

for r_idx, row in enumerate(lookup_data, start=2):
    ws_lookup.append(row)
    for col_idx in range(1, len(lookup_headers) + 1):
        c = ws_lookup.cell(row=r_idx, column=col_idx)
        c.font = cell_font
        c.border = border_grid
        if r_idx % 2 == 1:
            c.fill = zebra_stripe

for col in ws_lookup.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws_lookup.column_dimensions[col_letter].width = max(max_len * 1.4, 15)

# Write Solutions Sheet
sol_headers = ["情境模組", "欄位清洗目標", "推薦操作功能", "核心推薦公式", "解題關鍵技巧說明"]
sol_data = [
    ["A.訂艙清單", "提單號碼去空格與轉大寫", "資料剖析 / 快速填入", "=UPPER(TRIM(C2))", "去除頭尾多餘空格，強制轉換為大寫英文碼，消除隱藏字元。"],
    ["A.訂艙清單", "貨櫃號碼連字號去除", "尋找與取代 (Ctrl+H)", "=SUBSTITUTE(D2, \"-\", \"\")", "取代掉中線連字號或空白，確保11碼ISO貨櫃編號格式一致。"],
    ["A.訂艙清單", "重量/材積純數值萃取", "文字給資料行 / 替代函數", "=VALUE(SUBSTITUTE(SUBSTITUTE(I2,\"KGS\",\"\"),\",\",\"\"))", "將夾雜單位文字與千分位的字串轉為可運算的真實數值。"],
    ["A.訂艙清單", "日期規格化為標準西元日", "資料剖析 / DATE函數", "=DATE(LEFT(L2,4), MID(L2,6,2), RIGHT(L2,2))", "統一斜線、破折號與小數點分隔符號，格式化為YYYY-MM-DD。"],
    ["B.延滯費D&D", "實際在場天數計算", "數值運算與IF防呆", "=IF(ISBLANK(M2), TODAY()-L2, M2-L2)", "若尚未還櫃(空值)則計算至今日天數；若已還櫃則計算兩日之差。"],
    ["B.延滯費D&D", "逾期收費天數判定", "MAX函數邏輯判斷", "=MAX(0, 實際天數 - 免費天數)", "使用MAX避免產生負數天數，低於免收期自動歸零。"],
    ["B.延滯費D&D", "滯期金額計算與貨櫃加乘", "XLOOKUP / VLOOKUP", "=逾期天數 * XLOOKUP(櫃型, 費率表, 費率)", "動態依據20呎或40呎不同費率乘上逾期天數，計算應收帳款。"],
    ["C.帳單清理", "外幣幣別代碼分離", "LEFT / MID / TRIM", "=TRIM(LEFT(O2, 3))", "提取前3碼幣別字串（如USD、EUR、TWD），方便後續分組匯總。"],
    ["C.帳單清理", "原幣金額純淨轉型", "SUBSTITUTE / VALUE", "=VALUE(MID(O2, 4, 100))", "去除符號與代碼後將金額轉為數值，作為匯率換算乘數。"],
    ["C.帳單清理", "折合新台幣與5%營業稅", "四捨五入 ROUND函數", "=ROUND(原幣金額 * 匯率, 0)", "依財務會計準則折合TWD，並計算5%稅額與含稅總額。"],
    ["D.港口船期", "港口代碼與名稱對照標準化", "XLOOKUP函數", "=XLOOKUP(H2, 港口表!A:A, 港口表!B:B, \"查無港口\")", "透過全球UN/LOCODE主檔自動比對並標準化港口名稱與代碼。"],
    ["D.港口船期", "航行天數與跨月計算", "DATEDIF / 簡單減法", "=N2 - M2", "計算ETD至ETA之實際航行時間(Transit Time)，供航線績效分析。"]
]
ws_keys.append(sol_headers)
for col_idx in range(1, len(sol_headers) + 1):
    c = ws_keys.cell(row=1, column=col_idx)
    c.fill = PatternFill(start_color="783F04", end_color="783F04", fill_type="solid")
    c.font = header_font_white
    c.alignment = Alignment(horizontal="center", vertical="center")

for r_idx, row in enumerate(sol_data, start=2):
    ws_keys.append(row)
    for col_idx in range(1, len(sol_headers) + 1):
        c = ws_keys.cell(row=r_idx, column=col_idx)
        c.font = cell_font
        c.border = border_grid
        if r_idx % 2 == 1:
            c.fill = zebra_stripe

for col in ws_keys.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws_keys.column_dimensions[col_letter].width = max(max_len * 1.3, 16)

ex_file_path = "c:\\proj\\w12\\航運行政Excel資料前處理_綜合練習題庫_300筆.xlsx"
wb_ex.save(ex_file_path)
print(f"Saved exercise file to: {ex_file_path}")
