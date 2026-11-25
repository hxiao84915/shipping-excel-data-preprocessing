import os
import random
import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Set fixed seed for reproducibility
random.seed(42)

def create_30_demo_data():
    raw_data = [
        # (No, BL_No, Vessel_Voyage, POL, POD, Container_No, Cntr_Type, Pkgs, Weight, CBM, CutOff, ETD, Terms, Freight, Shipper)
        (1, "  EMCU20491823 ", "EVER GIVEN / 0124-055W", "TWKHH ", "USLAX", "EMCU8273641", "40'HC", "1,200 CTNS", " 18,450.50 KGS ", " 58.20 CBM ", "2026/11/02", "2026/11/05", "PREPAID", "USD 2,450.00", "  GREAT OCEAN TRADING CO.  "),
        (2, "OOCLU918273645", "OOCL ROTTERDAM / 088E", "tw tpe", "DEHAM ", "OOLU1928374", "40HQ", "850 PKGS", "12400", "42.5", "2026-11-03", "2026-11-07", "collect", "1,850", "Apex Logistics Corp."),
        (3, "MSCU992817263", "MSC ISABELLA / 2311W", "twkhh", " uslgB", "MSCU3847291", "20GP", "500 PLTS", " 9,800.00 KGS", " 28.00 CBM ", "2026.11.04", "2026.11.08", " Prepaid ", "$1,250.00", "FORMOSA PLASTICS CORP."),
        (4, "YMLU019283746", "YM WELLHEAD / 221E", "TWTXG", "JPTYO", "YMLU5566778", "20'GP", "320 CTNS", "5600 KGS", "18.4 CBM", "2026/11/05", "2026/11/08", "PREPAID", "USD 950.00", "CHENG SHIN RUBBER IND."),
        (5, "  WHLU883726152", "WAN HAI 315 / W082", "TWKEL ", " SGSIN ", "WHLU9900112", "40 DRY", "1,500 CTNS", "21,300.00", "62.0", "2026/11/06", "2026-11-10", "COLLECT", " 2,100.00 ", "KING CAR FOOD IND. CO."),
        (6, "CMAU771829304", "CMA CGM ANTOINE / 102W", "twkhh", "FRFOS", "CMAU4455667", "40HQ", "980 PKGS", " 16,200.00 KGS ", "52.8 CBM", "2026-11-06", "2026/11/11", "Prepaid", "USD 2,800", "uni-president enterprises"),
        (7, "ONEU662718293", "ONE CONTINUITY / 045E", "TWTPE", "USOAK", "ONEU1122334", "40HC", "1,100 CTNS", "19,500.00 KGS", "59.0 CBM", "2026.11.07", "2026.11.12", "COLLECT", "$2,650.00", " Giant Manufacturing Co. "),
        (8, "EMCU551928374", "EVER GLORY / 0118-032E", "TWKHH", "NLRTM", "EMCU9988776", "20GP", "450 PLTS", "8,950.00", "26.5", "2026/11/08", "2026/11/13", "PREPAID", " 1,350.00 ", "DELTA ELECTRONICS INC."),
        (9, "HLCU441928301", "AL DAHNA EXPRESS / 022W", "twtxg", "AEJEA", "HLCU3344556", "40HQ", "1,350 CTNS", "20,100.00 KGS", "61.5 CBM", "2026-11-09", "2026-11-14", "collect", "USD 3,100.00", "TATUNG COMPANY LTD"),
        (10, "SITC331928374", "SITC BANGKOK / 2209S", "TWKEL", "VNSGN", "SITU2233445", "20'GP", "600 CTNS", " 7,800.00 KGS ", " 22.00 CBM ", "2026.11.10", "2026.11.15", "Prepaid", "$880.00", "TAIWAN CEMENT CORP."),
        (11, "  OOCLU882736192", "OOCL POLAND / 073W", "tw tpe", "GBFLX", "OOLU6677889", "40'HC", "1,050 PKGS", "17,800.00", "55.0", "2026/11/11", "2026/11/16", "COLLECT", " 2,750.00 ", "ACER INCORPORATED"),
        (12, "YMLU991827364", "YM WITNESS / 185E", "TWKHH", "USSEA", "YMLU1133557", "40HQ", "1,280 CTNS", "22,400.00 KGS", "63.2 CBM", "2026-11-12", "2026-11-17", "PREPAID", "USD 2,900.00", "  ASUSTEK COMPUTER INC.  "),
        (13, "MSCU772819203", "MSC GULSUN / 2305W", "twkhh", "ITGOA", "MSCU9911223", "20GP", "480 PLTS", "9,200.00 KGS", "27.5 CBM", "2026.11.12", "2026.11.18", "collect", "$1,420.00", "NAN YA PLASTICS CORP."),
        (14, "WHLU661928374", "WAN HAI 502 / W103", "TWKEL", "MYPKG", "WHLU4466880", "40 DRY", "1,400 CTNS", "19,800.00", "58.0", "2026/11/13", "2026/11/19", "Prepaid", " 1,750.00 ", "QUANTA COMPUTER INC."),
        (15, "CMAU552819203", "CMA CGM BENJAMIN / 089W", "TWTXG", "ESVLC", "CMAU2244668", "40HC", "920 PKGS", " 15,600.00 KGS ", "49.0 CBM", "2026-11-14", "2026-11-20", "COLLECT", "USD 2,600.00", "compal electronics inc"),
        (16, "EMCU442819203", "EVER BRAVE / 0214-041E", "TWKHH", "USLAX", "EMCU3355779", "40'HC", "1,150 CTNS", "18,900.00 KGS", "57.5 CBM", "2026.11.15", "2026.11.21", "PREPAID", "$2,500.00", "AU OPTRONICS CORP."),
        (17, "ONEU332819203", "ONE HONG KONG / 077E", "tw tpe", "CAPRR", "ONEU8800224", "20GP", "550 CTNS", "8,400.00", "24.0", "2026/11/16", "2026/11/22", "collect", " 1,600.00 ", "INVENTEC CORPORATION"),
        (18, "HLCU221928304", "HAMBURG EXPRESS / 034W", "twkhh", "BEZEE", "HLCU5577991", "40HQ", "1,220 PKGS", " 19,300.00 KGS ", "58.8 CBM", "2026-11-17", "2026-11-23", "Prepaid", "USD 2,850.00", "WISTRON CORPORATION"),
        (19, "SITC112839402", "SITC SHANGHAI / 2211N", "TWKEL", "CNSHA", "SITU6688002", "20'GP", "700 CTNS", "10,500.00 KGS", "30.0 CBM", "2026.11.18", "2026.11.23", "COLLECT", "$650.00", "pegatron corporation"),
        (20, "OOCLU771928304", "OOCL GERMANY / 064E", "TWTXG", "USOAK", "OOLU3355771", "40HC", "1,300 CTNS", "21,000.00", "61.0", "2026/11/19", "2026/11/24", "PREPAID", " 2,700.00 ", "MEDIATEK INC."),
        (21, "  YMLU882719203 ", "YM WISH / 092E", "TWKHH ", "USLGB ", "YMLU7799113", "40'HC", "1,180 CTNS", " 19,100.00 KGS ", " 56.40 CBM ", "2026-11-20", "2026-11-25", "collect", "USD 2,550.00", "  FOXCONN TECHNOLOGY GROUP  "),
        (22, "MSCU661928304", "MSC LORRAINE / 2315E", "tw tpe", "AUMEL", "MSCU2255881", "20GP", "520 PLTS", "9,400.00", "28.2", "2026.11.21", "2026.11.26", "Prepaid", "$1,550.00", "Far Eastern New Century"),
        (23, "WHLU551928304", "WAN HAI 288 / S014", "TWKEL", "THBKK", "WHLU1144770", "40 DRY", "1,450 CTNS", "20,500.00 KGS", "60.0 CBM", "2026/11/22", "2026/11/27", "COLLECT", " 1,450.00 ", "EVERGREEN STEEL CORP."),
        (24, "CMAU441928304", "CMA CGM MARCO POLO / 115W", "twkhh", "GRPIR", "CMAU8811447", "40HQ", "1,080 PKGS", "18,200.00 KGS", "54.0 CBM", "2026-11-23", "2026-11-28", "PREPAID", "USD 2,950.00", "ADVANCED SEMICONDUCTOR ENGR"),
        (25, "ONEU221928304", "ONE HARBOUR / 052W", "TWTXG", "SGSIN", "ONEU4477003", "20'GP", "620 CTNS", " 8,100.00 KGS ", " 23.50 CBM ", "2026.11.24", "2026.11.29", "collect", "$920.00", "TSMC LOGISTICS DIVISION"),
        (26, "HLCU111928304", "PARIS EXPRESS / 019E", "tw tpe", "USNYC", "HLCU7700336", "40HC", "1,260 CTNS", "21,800.00", "62.8", "2026/11/25", "2026/11/30", "Prepaid", " 3,400.00 ", "LITE-ON TECHNOLOGY CORP."),
        (27, "EMCU331928304", "EVER FORTUNE / 0142-028W", "TWKHH", "TRIZM", "EMCU6699225", "40'HC", "950 PKGS", "16,800.00 KGS", "51.2 CBM", "2026-11-26", "2026-12-01", "COLLECT", "USD 3,250.00", "  SYNOLOGY INC.  "),
        # Row 28 is a direct duplicate of Row 4 (to test Remove Duplicates)
        (28, "YMLU019283746", "YM WELLHEAD / 221E", "TWTXG", "JPTYO", "YMLU5566778", "20'GP", "320 CTNS", "5600 KGS", "18.4 CBM", "2026/11/05", "2026/11/08", "PREPAID", "USD 950.00", "CHENG SHIN RUBBER IND."),
        (29, "SITC001928304", "SITC INCHON / 2215N", "TWKEL", "KRINC", "SITU3366992", "20GP", "490 PLTS", " 8,650.00 KGS ", " 25.00 CBM ", "2026.11.27", "2026.12.02", "PREPAID", "$780.00", "HIWIN TECHNOLOGIES CORP."),
        # Row 30 has missing/faulty entries
        (30, "OOCLU661928304", "OOCL ASIA / 140W", "twkhh", "USLAX", "OOLU9922558", "40HQ", "1,120 CTNS", "18,500.00", "56.0", "2026/11/28", "2026/12/03", "collect", " 2,480.00 ", "realtek semiconductor corp")
    ]
    
    cols = [
        "項次", "提單號碼_BL_NO", "船名航次_VESSEL_VOY", "起運港_POL", "目的港_POD",
        "貨櫃號碼_CNTR_NO", "櫃型_TYPE", "包裝件數_PKGS", "貨物毛重_KGS", "貨物材積_CBM",
        "結關日_CUTOFF", "開航日_ETD", "運費條款_TERMS", "運費金額_FREIGHT", "託運人_SHIPPER"
    ]
    
    df_raw = pd.DataFrame(raw_data, columns=cols)
    return df_raw

def clean_30_demo_data(df_raw):
    # Perform standard cleaning
    df = df_raw.copy()
    
    # 1. Deduplicate by 提單號碼
    df['提單號碼_CLEAN'] = df['提單號碼_BL_NO'].astype(str).str.strip().str.upper()
    df = df.drop_duplicates(subset=['提單號碼_CLEAN'], keep='first').copy()
    
    # 2. Text standardization
    df['提單號碼'] = df['提單號碼_CLEAN']
    
    # Split Vessel and Voyage
    split_vessel = df['船名航次_VESSEL_VOY'].str.split('/', expand=True)
    df['船名'] = split_vessel[0].str.strip()
    df['航次'] = split_vessel[1].str.strip() if split_vessel.shape[1] > 1 else ''
    
    # POL / POD standardize
    df['起運港'] = df['起運港_POL'].astype(str).str.strip().str.upper().str.replace(' ', '')
    df['目的港'] = df['目的港_POD'].astype(str).str.strip().str.upper().str.replace(' ', '')
    
    # Container No
    df['貨櫃號碼'] = df['貨櫃號碼_CNTR_NO'].astype(str).str.strip().str.upper().str.replace('-', '').str.replace(' ', '')
    
    # Container Type standardize (20GP, 40GP, 40HQ)
    def clean_type(val):
        v = str(val).upper().replace("'", "").replace(" ", "")
        if "40HC" in v or "40HQ" in v:
            return "40HQ"
        elif "20GP" in v:
            return "20GP"
        elif "40DRY" in v or "40GP" in v:
            return "40GP"
        return v
    df['標準櫃型'] = df['櫃型_TYPE'].apply(clean_type)
    
    # Pkgs extraction
    df['件數數值'] = df['包裝件數_PKGS'].astype(str).str.replace(',', '').str.extract(r'(\d+)')[0].astype(int)
    def get_unit(val):
        v = str(val).upper()
        for u in ['CTNS', 'PKGS', 'PLTS']:
            if u in v:
                return u
        return 'CTNS'
    df['包裝單位'] = df['包裝件數_PKGS'].apply(get_unit)
    
    # Weight KGS
    df['毛重_KGS'] = df['貨物毛重_KGS'].astype(str).str.replace(',', '').str.replace('KGS', '', case=False).str.strip().astype(float)
    
    # CBM
    df['材積_CBM'] = df['貨物材積_CBM'].astype(str).str.replace('CBM', '', case=False).str.strip().astype(float)
    
    # Dates to standard YYYY-MM-DD
    def clean_date(val):
        s = str(val).strip().replace('.', '/').replace('-', '/')
        parts = s.split('/')
        if len(parts) == 3:
            y, m, d = parts[0], parts[1], parts[2]
            if len(y) == 2:
                y = "20" + y
            return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"
        return s
    df['結關日'] = df['結關日_CUTOFF'].apply(clean_date)
    df['開航日'] = df['開航日_ETD'].apply(clean_date)
    
    # Freight Terms
    df['運費條款'] = df['運費條款_TERMS'].astype(str).str.strip().str.upper()
    
    # Freight Amount
    def clean_money(val):
        s = str(val).replace('USD', '').replace('$', '').replace(',', '').strip()
        return float(s)
    df['運費_USD'] = df['運費金額_FREIGHT'].apply(clean_money)
    
    # Shipper
    df['託運人'] = df['託運人_SHIPPER'].astype(str).str.strip().str.upper()
    
    clean_cols = [
        "提單號碼", "船名", "航次", "起運港", "目的港",
        "貨櫃號碼", "標準櫃型", "件數數值", "包裝單位",
        "毛重_KGS", "材積_CBM", "結關日", "開航日",
        "運費條款", "運費_USD", "託運人"
    ]
    df_clean = df[clean_cols].reset_index(drop=True)
    df_clean.insert(0, "項次", range(1, len(df_clean) + 1))
    return df_clean

print("Generating 30 demo records...")
df_demo_raw = create_30_demo_data()
df_demo_clean = clean_30_demo_data(df_demo_raw)
print(f"Demo raw shape: {df_demo_raw.shape}, clean shape: {df_demo_clean.shape}")

# Save to Excel Demo
wb_demo = openpyxl.Workbook()
ws_raw = wb_demo.active
ws_raw.title = "01_原始待處理資料(Dirty)"
ws_clean = wb_demo.create_sheet(title="02_前處理完成資料(Cleaned)")

# Style helpers
header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
header_font = Font(name="Microsoft JhengHei", size=11, bold=True, color="FFFFFF")
data_font = Font(name="Microsoft JhengHei", size=10)
border_thin = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)
zebra_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")

def write_sheet(ws, df, is_raw=True):
    # Write headers
    headers = list(df.columns)
    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    # Write rows
    for r_idx, row in enumerate(df.values, start=2):
        ws.append(list(row))
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=r_idx, column=col_idx)
            cell.font = data_font
            cell.border = border_thin
            if r_idx % 2 == 1:
                cell.fill = zebra_fill
            if col_idx == 1:
                cell.alignment = Alignment(horizontal="center")
    
    # Auto column width
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len * 1.3, 12)

write_sheet(ws_raw, df_demo_raw, is_raw=True)
write_sheet(ws_clean, df_demo_clean, is_raw=False)

demo_file_path = "c:\\proj\\w12\\航運行政Excel資料前處理_示範資料檔_30筆.xlsx"
wb_demo.save(demo_file_path)
print(f"Saved demo file to: {demo_file_path}")
