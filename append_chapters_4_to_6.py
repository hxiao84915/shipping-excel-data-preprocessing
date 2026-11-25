import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

from build_full_manual import (
    doc, add_p, add_h1, add_h2, add_h3, add_callout, add_figure, add_table_data,
    COLOR_PRIMARY, COLOR_SECONDARY, COLOR_ACCENT, COLOR_DARK, COLOR_MUTED
)
from build_exercise_data import df_exercise_300

fig_dir = "c:\\proj\\w12\\figures"

print("Writing Chapter 4...")
# ===========================================================================
# CHAPTER 4
# ===========================================================================
add_h1(doc, "第四章 前處理必備功能選項與核心函數公式速查寶典")

add_h2(doc, "4.1 功能區選單操作手冊 (資料剖析、移除重複、閃電填入、資料驗證)")
add_p(doc, "Excel 在『資料』(Data) 索引標籤中提供了一系列功能強大的前處理命令。對於初學者而言，善用這些內建工具往往不需要撰寫複雜公式即可在數秒內完成龐大數據的清洗。以下針對四項最重要的選單功能進行深度實戰操作說明：")

add_h3(doc, "4.1.1 資料剖析 (Text to Columns)")
add_p(doc, "【核心功能】：將存放在單一儲存格中的複合長字串，依據特定的『分隔符號』（如逗號、Tab、空格、斜線、破折號）或『固定寬度』精確切分至多個相鄰的直欄。")
add_p(doc, "【航運實戰場景】：")
add_p(doc, "1. 拆解『船名 / 航次』：以斜線 \"/\" 作為其他分隔符號切分。")
add_p(doc, "2. 拆解『包裝件數與單位』：以半形空格 \" \" 作為分隔符號，將 \"1,500 CTNS\" 拆解為件數數值欄與單位名稱欄。")
add_p(doc, "3. 修正非標準文字型日期：在精靈步驟 3/3 中，選取欄位格式為【日期 (YMD)】，Excel 會強制將 \"20261105\" 或 \"2026.11.05\" 等頑固字串轉換為底層可運算的西元序列日期。")

add_h3(doc, "4.1.2 移除重複項 (Remove Duplicates)")
add_p(doc, "【核心功能】：快速掃描資料表，辨識並刪除具有相同值的重複資料列，僅保留第一筆出現的唯一紀錄。")
add_p(doc, "【航運實戰防呆秘訣】：")
add_p(doc, "在彈出的對話方塊中，『絕不可全選所有欄位』！若全選所有欄位，只有當該列的所有資訊（包括備註與運費）完全一模一樣時才會被判定為重複。在航運實務中，不同人員匯入時可能備註稍有不同，因此必須『單獨勾選業務唯一主鍵』，例如：")
add_p(doc, "● 海運出口主檔：僅勾選【提單號碼 (B/L No.)】或【貨櫃號碼 (Container No.)】。")
add_p(doc, "● 貨代會計帳單：僅勾選【帳單編號 (Invoice No.)】或【請款項次編號】。")

add_h3(doc, "4.1.3 閃電填入 / 快速填入 (Flash Fill / 快速鍵 Ctrl+E)")
add_p(doc, "【核心功能】：Excel 365 / 2016 以上版本具備的機器學習智慧字串模式辨識技術。使用者只需在相鄰直欄手動鍵入 1 到 2 個期望的輸出結果範例，按下 Ctrl+E，系統便會自動推論切分或組合邏輯，瞬間填滿整欄數千列資料。")
add_p(doc, "【適用情境與限制】：極度適合用於規則明確的文字擷取（如從櫃號前 4 碼提取船東英文代碼、從託運人全名中提取核心公司名稱）。但需特別注意：若原始資料存在格式不一致的例外（例如部分航次帶有方向代碼 E/W，部分沒有），閃電填入可能會發生少數推論偏差，填入後務必透過篩選檢查最後幾列。")

add_h3(doc, "4.1.4 資料驗證 (Data Validation) 下拉防呆選單")
add_p(doc, "【核心功能】：在資料輸入階段直接從源頭杜絕髒資料產生，限定儲存格僅能輸入特定範圍的數值、日期，或強制使用者由下拉選單中挑選合法項目。設定畫面如圖 4-4 所示：")

add_figure(doc, os.path.join(fig_dir, "fig7_data_validation_dialog.png"), "圖 4-4  Excel 資料驗證 (Data Validation) 貨櫃規格下拉防呆設定")

add_p(doc, "【航運實務設定三步驟】：")
add_p(doc, "1. 選取櫃型整欄（例如 G2:G1000）。")
add_p(doc, "2. 點擊【資料】->【資料驗證】，在【儲存格內允許】下拉選單選取【清單】(List)。")
add_p(doc, "3. 在【來源】輸入框中鍵入標準規格：20GP, 40GP, 40HQ, 45HQ, 20RF, 40RF（項目間以半形逗點區隔）。按下確定後，儲存格右側即會出現下拉箭頭，徹底根除同仁隨興手打 40'HC、40DC 等錯誤名稱！")

add_h2(doc, "4.2 航運資料前處理「十大核心函數」語法與精選範例")
add_p(doc, "當面對複雜或動態連結的航運數據時，靈活組合公式函數能達到 100% 精準與自動化的前處理效果。圖 4-5 展示了航運前處理四大體系的十大核心函數架構：")

add_figure(doc, os.path.join(fig_dir, "fig8_core_functions_architecture.png"), "圖 4-5  航運資料前處理『十大核心函數』體系架構導圖")

add_p(doc, "以下詳細解析各函數的標準語法、參數意義與航運實務範例（表 4-1 提供速查總覽）：")

func_headers = ["函數名稱", "語法結構", "功能分類", "航運實務典型應用範例", "範例預期運算結果"]
func_rows = [
    ["TRIM", "=TRIM(text)", "文字清洗", "=TRIM(\"  EMCU1234567  \")", "\"EMCU1234567\" (去除前後空白)"],
    ["CLEAN", "=CLEAN(text)", "文字清洗", "=CLEAN(A2)", "清除系統匯出的不可列印換行ASCII碼"],
    ["UPPER", "=UPPER(text)", "大小寫轉換", "=UPPER(\"twkhh\")", "\"TWKHH\" (統一港口代碼全大寫)"],
    ["LEFT", "=LEFT(text, [num])", "字串擷取", "=LEFT(\"EMCU8273641\", 4)", "\"EMCU\" (擷取貨櫃箱主4碼英文字頭)"],
    ["MID", "=MID(text, start, num)", "字串擷取", "=MID(\"2026-11-05\", 6, 2)", "\"11\" (精準擷取日期中的月份數字)"],
    ["RIGHT", "=RIGHT(text, [num])", "字串擷取", "=RIGHT(\"0124-055W\", 1)", "\"W\" (擷取航次方向西向指示碼)"],
    ["FIND", "=FIND(find_text, within)", "位置定位", "=FIND(\"/\", \"EVER GIVEN / 055W\")", "12 (回傳斜線在字串中的第12個位置)"],
    ["SUBSTITUTE", "=SUBSTITUTE(text, old, new)", "字串替換", "=SUBSTITUTE(\"15,000 KGS\", \"KGS\", \"\")", "\"15,000 \" (去除重量文字單位)"],
    ["VALUE", "=VALUE(text)", "型態轉換", "=VALUE(\"2450.50\")", "2450.5 (文字數字轉為真數值可供計算)"],
    ["DATE", "=DATE(year, month, day)", "日期構建", "=DATE(2026, 11, 5)", "2026/11/05 (標準西元日期序號值)"],
    ["XLOOKUP", "=XLOOKUP(key, lk_col, ret_col)", "比對查找", "=XLOOKUP(\"TWKHH\", 港口!A:A, 港口!B:B)", "\"台灣高雄港\" (跨表秒速比對港口主檔)"],
    ["IFERROR", "=IFERROR(value, val_if_err)", "防呆容錯", "=IFERROR(VLOOKUP(...), \"待確認\")", "當查詢查無資料時顯示友善自訂提示"]
]
add_table_data(doc, func_headers, func_rows, col_widths=[1.1, 1.6, 1.0, 1.8, 1.5], caption="表 4-1  前處理十大核心函數語法、用途與航運實務範例總覽")

add_h3(doc, "4.2.1 字串清洗家族：TRIM 與 CLEAN")
add_p(doc, "● TRIM(text)：")
add_p(doc, "  - 說明：專門清除字串開頭與尾端的所有半形空格，若字串中間有連續多個空格，則會自動合併為單一空格。這是解決 VLOOKUP 比對失敗的頭號功臣。")
add_p(doc, "● CLEAN(text)：")
add_p(doc, "  - 說明：從舊版海運 AS400 主機或外部 EDI 檔案匯出的報表中，文字尾端常夾帶隱藏的換行符號（ASCII 碼 10 或 13），肉眼完全看不出但會讓公式判定為不等於。搭配 =CLEAN(TRIM(A2)) 可徹底消除一切頑固字元。")

add_h3(doc, "4.2.2 英文規範家族：UPPER、LOWER 與 PROPER")
add_p(doc, "● UPPER(text)：強制轉換為全大寫。航運代碼（提單號、港口 UN/LOCODE、櫃號）一律必須使用 UPPER 標準化。")
add_p(doc, "● LOWER(text)：強制轉換為全小寫。可用於客戶電子郵件（Email）標準化。")
add_p(doc, "● PROPER(text)：字首大寫、其餘小寫。極度適用於整理外國客戶人名或船東全稱（例如將 \"john smith\" 轉換為 \"John Smith\"）。")

add_h3(doc, "4.2.3 文字定位與擷取：LEFT、MID、RIGHT、LEN 與 FIND")
add_p(doc, "● LEFT(text, num_chars)：從文字左側開始截取指定長度的字元。")
add_p(doc, "● RIGHT(text, num_chars)：從文字右側截取指定長度的字元。")
add_p(doc, "● MID(text, start_num, num_chars)：從指定起始位置截取特定長度字元。")
add_p(doc, "● FIND(find_text, within_text)：尋找特定字元在字串中出現的第一個位置（區分大小寫）。")
add_p(doc, "● LEN(text)：計算文字字串的總字元長度（含空格）。")

add_h3(doc, "4.2.4 型態轉換與替代：SUBSTITUTE 與 VALUE")
add_p(doc, "● SUBSTITUTE(text, old_text, new_text, [instance_num])：")
add_p(doc, "  - 將字串中的特定文字替換為新文字。例如將連字號去除：SUBSTITUTE(A2, \"-\", \"\")。")
add_p(doc, "● VALUE(text)：")
add_p(doc, "  - 將帶有引號、靠左對齊的文字型態數字，強制轉換為可運算的標準數值。若儲存格內殘留有字母或特殊符號，VALUE 會拋出 #VALUE!，因此必須先用 SUBSTITUTE 把文字清乾淨後再餵給 VALUE。")

add_h3(doc, "4.2.5 查詢對照與防呆：XLOOKUP 與 IFERROR")
add_p(doc, "● XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode])：")
add_p(doc, "  - Excel 365 / 2021 最強大革命性查詢函數，全面取代傳統 VLOOKUP。XLOOKUP 支援向左查詢、無須計算第幾欄、預設完全比對，並內建查無資料時的預設值，不再輕易爆出 #N/A。")
add_p(doc, "  - 範例：=XLOOKUP(C2, 港口代碼表!A:A, 港口代碼表!B:B, \"非標準代碼\")")
add_p(doc, "● IFERROR(value, value_if_error)：")
add_p(doc, "  - 攔截任何計算錯誤並優雅替換。例如計算櫃量比例時避免除以 0：=IFERROR(A2/B2, 0)。")

add_h2(doc, "4.3 航運業務複合巢狀公式進階解題示範")
add_p(doc, "在實際海運工作場合，單一函數往往無法解決複雜的髒資料，必須將多個函數如同積木般層層嵌套。以下展示三個最高頻的複合巢狀公式模板：")

add_p(doc, "● 模板一：動態依斜線拆分船名與航次（無須手動指定長度）")
add_p(doc, "  - 提取船名（斜線左側）：")
add_p(doc, "    =TRIM(LEFT(A2, FIND(\"/\", A2&\"/\") - 1))", font_size=9.5, bold=True, color=COLOR_PRIMARY)
add_p(doc, "  - 提取航次（斜線右側）：")
add_p(doc, "    =TRIM(MID(A2, FIND(\"/\", A2&\"/\") + 1, 100))", font_size=9.5, bold=True, color=COLOR_PRIMARY)
add_p(doc, "  - 技巧亮點：在搜尋目標加上 &\"/\" 防呆技巧，即使原始資料忘記打斜線，公式也不會崩潰報錯！")

add_p(doc, "● 模板二：去除千分位逗點與計量單位，轉化為純數值加總")
add_p(doc, "  - 處理類似 \" 24,500.50 KGS \" 或 \"$1,850.00\"：")
add_p(doc, "    =VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(A2), \"KGS\", \"\"), \"$\", \"\"), \",\", \"\"))", font_size=9.5, bold=True, color=COLOR_PRIMARY)

add_p(doc, "● 模板三：將各類點號/減號文字重構為標準西元日期")
add_p(doc, "  - 處理類似 \"2026.11.05\" 或 \"2026-11-05\"：")
add_p(doc, "    =DATE(LEFT(A2, 4), MID(SUBSTITUTE(SUBSTITUTE(A2, \".\", \"/\"), \"-\", \"/\"), 6, 2), RIGHT(A2, 2))", font_size=9.5, bold=True, color=COLOR_PRIMARY)

# ===========================================================================
# CHAPTER 5
# ===========================================================================
add_h1(doc, "第五章 300筆航運情境實戰練習題庫與解題全攻略")

add_h2(doc, "5.1 300筆實戰題庫四大業務情境架構導覽")
add_p(doc, "為了讓初學者在學習完基礎操作與函數後能夠獲得充足的實戰演練，本手冊精心建構了一份涵蓋四大航運實務領域的『300筆情境練習題庫』（隨附於『航運行政Excel資料前處理_綜合練習題庫_300筆.xlsx』）。四大模組結構如圖 5-1 所示：")

add_figure(doc, os.path.join(fig_dir, "fig10_scenario_300_overview.png"), "圖 5-1  300筆航運情境練習題庫四大業務模組架構圖")

add_p(doc, "四大情境模組分別對應航運行政人員職涯中最常面臨的資料痛點：")
add_p(doc, "● 模組 A（第 1 - 100 筆）：海運出口訂艙與裝船清單清理（聚焦提單、櫃號、件數重量原子化）。")
add_p(doc, "● 模組 B（第 101 - 170 筆）：貨櫃滯期延滯費 (D&D) 計費天數與逾期分析（聚焦日期相減、空值未還櫃判斷、MAX 截斷與動態費率乘算）。")
add_p(doc, "● 模組 C（第 171 - 240 筆）：國際貨代應收應付帳單幣別、匯率與金額清理（聚焦多外幣文字分離、匯率折合新台幣與 5% 營業稅精算）。")
add_p(doc, "● 模組 D（第 241 - 300 筆）：跨國港口代碼與船期表標準化（聚焦聯合國 UN/LOCODE 主檔對照、港口中文名稱查找與航行天數計算）。")

add_h2(doc, "5.2 模組 A：海運出口訂艙與裝船清單清理 (100筆)")
add_p(doc, "【情境背景】：你是某國際知名海運承攬公司（Freight Forwarder）的出口文件人員。船公司剛發送來本航次 100 筆裝船總清單，但資料是直接由舊系統匯出的未格式化文本，充滿多餘空白、小寫代碼、混雜單位的重量與不同格式的結關日。你必須將其整理為標準表，以利後續報關與製作出口艙單。")

# Preview Table 5-1 (First 10 rows of Module A)
mod_a_sample = df_exercise_300.iloc[0:10][["項次", "業務編號", "關聯識別碼", "規格項目", "主體名稱", "起點_POL", "迄點_POD", "量值1_件數", "量值2_重量", "起始基準日", "分類條款", "關係客戶"]].values
mod_a_headers = ["項次", "提單號碼(髒)", "貨櫃號碼(髒)", "櫃型", "船名航次", "POL", "POD", "件數(帶單位)", "毛重(帶單位)", "結關日", "運費條款", "託運人客戶"]
add_table_data(doc, mod_a_headers, [[str(x) for x in r] for r in mod_a_sample], caption="表 5-1  300筆實務題庫模組 A：出口訂艙裝船清單原始結構 (前10筆示範)")

add_p(doc, "【模組 A 解題步驟與核心公式全指引】：")
add_p(doc, "1. 提單號碼標準化：在相鄰直欄輸入公式 =UPPER(TRIM(C2))，去除首尾空白並轉為純大寫。")
add_p(doc, "2. 貨櫃號碼清洗：櫃號可能夾帶破折號（如 'OOLU-1928374'），輸入公式 =UPPER(SUBSTITUTE(TRIM(D2), \"-\", \"\"))，還原為 11 碼標準 ISO 櫃號。")
add_p(doc, "3. 件數純數值萃取：使用 =VALUE(LEFT(TRIM(I2), FIND(\" \", TRIM(I2)&\" \")-1)) 提取純包裝件數。")
add_p(doc, "4. 貨物毛重轉為浮點數：使用 =VALUE(SUBSTITUTE(SUBSTITUTE(UPPER(TRIM(J2)), \"KGS\", \"\"), \",\", \"\")) 提取可運算純重量。")
add_p(doc, "5. 結關日規格化：使用快速取代（Ctrl+H）將所有小數點 '.' 與減號 '-' 取代為斜線 '/'，並套用儲存格格式 YYYY-MM-DD。")
add_p(doc, "【成果驗收核對指標】：清洗完成後，選取 100 筆毛重欄位，狀態列加總（SUM）應能正確統計出總重量（約在 1,450,000 至 1,550,000 KGS 之間），且不能有任何一筆顯示 #VALUE! 錯誤！")

add_h2(doc, "5.3 模組 B：貨櫃滯期延滯費 (D&D) 計費天數與逾期分析前處理 (70筆)")
add_p(doc, "【情境背景】：在貨櫃海運運送中，貨櫃場站與船公司提供進口貨物一定天數的『免費使用期』(Free Time)。若貨主提領後未在免收期內還空櫃，將被課以昂貴的延滯費（Detention）。行政同仁必須根據『提櫃/卸船日』與『還空櫃日』計算實際在場天數，若貨櫃尚未歸還（還櫃日空白），則需計算至今日（TODAY），並精確計算逾期天數與帳單金額。")

# Preview Table 5-2 (Sample rows of Module B)
mod_b_sample = df_exercise_300.iloc[100:110][["項次", "業務編號", "關聯識別碼", "規格項目", "主體名稱", "量值1_件數", "量值2_重量", "起始基準日", "結束航行日", "分類條款", "費用金額", "關係客戶"]].values
mod_b_headers = ["項次", "帳單單號", "貨櫃號碼", "櫃型", "免費天數條件", "日租費率", "免收約定", "卸船提櫃日", "還空櫃日(含空值)", "計費類別", "每日單價", "託運人客戶"]
add_table_data(doc, mod_b_headers, [[str(x) for x in r] for r in mod_b_sample], caption="表 5-2  300筆實務題庫模組 B：貨櫃滯期延滯費 (D&D) 計算結構")

add_p(doc, "【模組 B 解題步驟與核心公式全指引】：")
add_p(doc, "1. 卸船日與還櫃日標準化：先確認起始基準日（卸船日）與結束基準日（還櫃日）已被正確轉為標準西元日期格式。")
add_p(doc, "2. 實際在場天數計算（含未還櫃防呆處理）：")
add_p(doc, "   - 若還櫃日為空值，代表貨主至今尚未還櫃，應計算自提櫃日起算至今（TODAY()）的在場天數。")
add_p(doc, "   - 公式：=IF(ISBLANK(I2), TODAY() - H2, I2 - H2)", font_size=9.5, bold=True, color=COLOR_PRIMARY)
add_p(doc, "3. 逾期計費天數判定（核心技巧：使用 MAX 避免負數）：")
add_p(doc, "   - 若在場天數為 5 天，約定免費天數為 7 天，相減為 -2 天。延滯費不能收負數，低於免收期應自動歸零。")
add_p(doc, "   - 公式：=MAX(0, 實際在場天數 - 免費天數)", font_size=9.5, bold=True, color=COLOR_PRIMARY)
add_p(doc, "4. 滯期費總金額精算：")
add_p(doc, "   - 依據櫃型（20呎每日 USD 50，40呎每日 USD 90），乘以逾期計費天數：")
add_p(doc, "   - 公式：=逾期天數 * IF(ISNUMBER(FIND(\"20\", D2)), 50, 90)")
add_p(doc, "【成果驗收核對指標】：逾期天數欄位不得出現任何負數天數；未還櫃之項目應能動態呈現最新累計天數。")

add_h2(doc, "5.4 模組 C：國際貨代應收應付帳單幣別、匯率與金額清理 (70筆)")
add_p(doc, "【情境背景】：貨運承攬業的財務結算極為複雜，單一出口航次涉及海運基本運費（USD）、港口吊櫃費 THC（TWD）、歐洲端報關費（EUR）、日本內陸卡車費（JPY）等多國幣別。行政同仁必須將混雜著貨幣符號與文字的金額字串拆解，分離幣別，乘上當期基準匯率折合新台幣（TWD），並依法精算 5% 營業稅額與含稅請款總額。")

# Preview Table 5-3 (Sample rows of Module C)
mod_c_sample = df_exercise_300.iloc[170:180][["項次", "業務編號", "關聯識別碼", "規格項目", "主體名稱", "量值1_件數", "量值2_重量", "起始基準日", "結束航行日", "分類條款", "費用金額", "關係客戶"]].values
mod_c_headers = ["項次", "請款單號", "提單號參照", "費用項目名稱", "匯率標準基準", "原幣別代碼", "折合匯率", "開立日期", "付款到期日", "收付類別", "原始金額字串", "往來客戶"]
add_table_data(doc, mod_c_headers, [[str(x) for x in r] for r in mod_c_sample], caption="表 5-3  300筆實務題庫模組 C：貨代帳單幣別匯率金額清理結構")

add_p(doc, "【模組 C 解題步驟與核心公式全指引】：")
add_p(doc, "1. 外幣幣別代碼提取：")
add_p(doc, "   - 觀察原始金額字串（如 \"USD 2,450.00\"、\"NT$ 15,000\"、\"JPY 80,000\"）。")
add_p(doc, "   - 公式：=IF(ISNUMBER(FIND(\"USD\", K2)), \"USD\", IF(ISNUMBER(FIND(\"NT\", K2)), \"TWD\", IF(ISNUMBER(FIND(\"EUR\", K2)), \"EUR\", IF(ISNUMBER(FIND(\"JPY\", K2)), \"JPY\", \"RMB\"))))")
add_p(doc, "2. 原幣金額純數值萃取：")
add_p(doc, "   - 消除幣別代碼、貨幣符號與逗點後轉數值：")
add_p(doc, "   - 公式：=VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(K2), \"USD\", \"\"), \"NT$\", \"\"), \"EUR\", \"\"), \"JPY\", \"\"), \"RMB\", \"\"), \",\", \"\"))", font_size=9, bold=True, color=COLOR_PRIMARY)
add_p(doc, "3. 折合新台幣未稅金額（依財務準則四捨五入至整數位）：")
add_p(doc, "   - 公式：=ROUND(原幣純金額 * 折合匯率, 0)")
add_p(doc, "4. 5% 營業稅額與含稅總額精算：")
add_p(doc, "   - 營業稅額：=ROUND(未稅TWD金額 * 0.05, 0)")
add_p(doc, "   - 請款含稅總額：=未稅TWD金額 + 營業稅額")
add_p(doc, "【成果驗收核對指標】：篩選【收付類別】，分別統計 AR（應收帳款）含稅總額與 AP（應付帳款）含稅總額，兩者相減即可得出該批貨物之承攬毛利（Gross Profit）。")

add_h2(doc, "5.5 模組 D：跨國港口代碼與船期表標準化 (60筆)")
add_p(doc, "【情境背景】：航線調配部門每月維護由各大船公司聯盟發布的國際船期班表（Vessel Schedule）。由於各船東對靠泊港口的簡稱不同（例如有的寫 Kaohsiung、有的寫 KHH、有的寫 TWKHH），行政人員必須將非標準的港口名稱與代碼，透過跨工作表對照表（Lookup Table）標準化為聯合國 5 碼 UN/LOCODE 國際代碼與標準中文站名，並計算 ETD 至 ETA 的實際航行天數（Transit Days）。")

# Preview Table 5-4 (Sample rows of Module D)
mod_d_sample = df_exercise_300.iloc[240:250][["項次", "業務編號", "關聯識別碼", "規格項目", "主體名稱", "起點_POL", "迄點_POD", "量值1_件數", "量值2_重量", "量值3_材積", "起始基準日", "結束航行日"]].values
mod_d_headers = ["項次", "船期編號", "待清洗港口碼", "航線代碼", "船名航次", "POL起運港", "POD目的港名稱", "所屬航運區域", "港口標準中文名", "標準UN/LOCODE", "預計抵達ETA", "預計開航ETD"]
add_table_data(doc, mod_d_headers, [[str(x) for x in r] for r in mod_d_sample], caption="表 5-4  300筆實務題庫模組 D：跨國港口與船期表標準化結構")

add_p(doc, "【模組 D 解題步驟與核心公式全指引】：")
add_p(doc, "1. 港口代碼去除空格與轉大寫：使用公式 =UPPER(TRIM(C2))。")
add_p(doc, "2. 跨表對照標準 UN/LOCODE 國際碼（使用 XLOOKUP 函數）：")
add_p(doc, "   - 在隨附活頁簿的『02_基礎代碼對照表(Lookup)』工作表中，記錄了標準港口清單。")
add_p(doc, "   - 依據目的港英文名稱查詢標準 UN 代碼：")
add_p(doc, "   - 公式：=XLOOKUP(G2, '02_基礎代碼對照表(Lookup)'!$A:$A, '02_基礎代碼對照表(Lookup)'!$B:$B, \"查無標準碼\")", font_size=9.5, bold=True, color=COLOR_PRIMARY)
add_p(doc, "3. 查找港口標準中文名稱：")
add_p(doc, "   - 公式：=XLOOKUP(G2, '02_基礎代碼對照表(Lookup)'!$A:$A, '02_基礎代碼對照表(Lookup)'!$C:$C, \"未知港口\")")
add_p(doc, "4. 靠泊作業與航行天數計算：")
add_p(doc, "   - 靠泊停留天數：=開航日(ETD) - 抵達日(ETA)。直接進行日期相減，格式設定為『通用』。")
add_p(doc, "【成果驗收核對指標】：所有靠泊港口皆成功對應為標準 5 碼大寫英文字元；航行停留天數介於 1 至 3 天之間，無負數或 #VALUE! 異常。")

add_h2(doc, "5.6 隨附活頁簿結構說明與練習驗收方式")
add_p(doc, "本手冊隨附的練習活頁簿『航運行政Excel資料前處理_綜合練習題庫_300筆.xlsx』內建三個功能分頁：")
add_p(doc, "1. 分頁一【01_300筆航運情境練習題庫(Raw)】：存放完整的 300 筆原始未清洗數據，供同仁親手操作清洗。")
add_p(doc, "2. 分頁二【02_基礎代碼對照表(Lookup)】：提供全球 14 大核心港口主檔，供模組 B 與模組 D 進行 XLOOKUP 參照練習。")
add_p(doc, "3. 分頁三【03_解題思路與公式指引(Solutions)】：彙整四大模組各欄位的推薦清洗公式、函數用法與業務操作要領，供學員完成練習後核對答案。")

# ===========================================================================
# CHAPTER 6
# ===========================================================================
add_h1(doc, "第六章 航運資料品質控管清單與參考文獻")

add_h2(doc, "6.1 航運行政人員資料前處理三階段 SOP 檢核表")
add_p(doc, "為確保航運數據在清洗過程中的完整性、正確性與業務合規性，建議行政同仁在日常作業中嚴格依照表 6-1 進行『三階段品質檢核』：")

chk_headers = ["檢核階段", "核心查核項目", "檢核要點與檢驗方法", "標準合格狀態", "確認勾選"]
chk_rows = [
    ["第一階段：清洗前 (Pre-check)", "1. 原始檔案備份存檔", "另存新檔並將工作表複製為獨立複本", "確認保留唯讀原始底稿", "[   ] 已完成"],
    ["第一階段：清洗前 (Pre-check)", "2. 標題列唯一性檢查", "檢查第 1 列是否有空白標題或合併儲存格", "每欄皆有獨立唯一的英文/中文標題", "[   ] 已完成"],
    ["第一階段：清洗前 (Pre-check)", "3. 總列數與總筆數記錄", "選取資料欄，透過狀態列記錄原始列數", "記錄初始筆數 (如 300 列)", "[   ] 已完成"],
    ["第二階段：清洗中 (In-process)", "4. 關鍵主鍵去空格大寫", "對提單號、櫃號套用 TRIM 與 UPPER", "徹底消除前後空白，字母全大寫", "[   ] 已完成"],
    ["第二階段：清洗中 (In-process)", "5. 欄位原子化分割檢查", "檢查船名航次、件數單位是否已切分獨立欄", "一欄僅含單一業務變數", "[   ] 已完成"],
    ["第二階段：清洗中 (In-process)", "6. 數值與日期型態校驗", "以 ISNUMBER 檢查重量金額與日期是否為真值", "所有數值靠右對齊且回傳 TRUE", "[   ] 已完成"],
    ["第三階段：清洗後 (Post-check)", "7. 重複記錄剔除確認", "執行移除重複項，核對剔除筆數與原因", "主鍵無任何重複", "[   ] 已完成"],
    ["第三階段：清洗後 (Post-check)", "8. 公式轉純值 (貼上值)", "全選清洗後欄位，複製 -> 選擇性貼上為『值』", "切斷外連公式，避免活頁簿運算延遲", "[   ] 已完成"],
    ["第三階段：清洗後 (Post-check)", "9. 業務邏輯合理性審查", "篩選檢查極端值（如毛重負數、ETD早於結關日）", "所有運費、天數、重量皆符合航運常理", "[   ] 已完成"]
]
add_table_data(doc, chk_headers, chk_rows, col_widths=[1.5, 1.4, 1.9, 1.5, 0.9], caption="表 6-1  航運行政人員資料前處理三階段 SOP 檢核查核表")

add_h2(doc, "6.2 參考文獻與國際海運標準規範註記")
add_p(doc, "本手冊之編製內容、標準代碼與操作規範，嚴格參照並引用以下國際航運標準組織與官方軟體技術文件（依需求註明參考來源）：")

add_p(doc, "1. 微軟官方技術文檔 (Microsoft Learn Documentation)")
add_p(doc, "   - Microsoft Corporation (2024). Excel Functions by Category: Text, Date and Time, Lookup and Reference Functions. 參照官方函數定義與運算規範。")
add_p(doc, "   - Microsoft Corporation (2024). Top Tips for Cleaning Your Data in Microsoft Excel. 官方資料前處理指南。")

add_p(doc, "2. 聯合國貿易便捷化與電子商務中心 (UN/CEFACT)")
add_p(doc, "   - United Nations Economic Commission for Europe (UNECE). UN/LOCODE: Codes for Trade and Transport Locations (Recommendation No. 16). 全球貿易與運輸港口5碼代碼國際標準。")

add_p(doc, "3. 國際標準化組織 (International Organization for Standardization, ISO)")
add_p(doc, "   - ISO 8601:2019. Data elements and interchange formats — Information interchange — Representation of dates and times. 規定國際標準西元日期表示法 (YYYY-MM-DD)。")
add_p(doc, "   - ISO 6346:2022. Freight containers — Coding, identification and marking. 規定貨櫃11碼箱號（4碼字頭+6碼流水號+1碼檢核碼）及尺寸型號代碼標準。")

add_p(doc, "4. 現代資料科學理論架構")
add_p(doc, "   - Hadley Wickham (2014). Tidy Data. Journal of Statistical Software, 59(10), 1-23. 本手冊第二章『乾淨資料』三大結構設計原理之學術淵源。")

add_p(doc, "5. 國際貨運代理協會聯合會 (FIATA)")
add_p(doc, "   - FIATA (2023). Multimodal Transport Guidelines and Standard Freight Forwarding Documentation. 海運提單（FBL）、訂艙確認件與費用結算標準規範。")

# ===========================================================================
# APPENDIX
# ===========================================================================
add_h1(doc, "附錄：實用代碼速查表與資源指南")

add_h2(doc, "附錄一：全球常用 50 大港口 UN/LOCODE 國際代碼速查表")
add_p(doc, "表 附錄-1 整理了全球海運最高頻之 50 大集裝箱港口的聯合國標準 5 碼 UN/LOCODE 代碼、英文正式名稱與中文稱謂，供行政同仁建立 XLOOKUP 比對主檔使用：")

port50_headers = ["港口名稱", "UN/LOCODE", "中文名稱", "所屬國家/地區", "主要所屬航線區域"]
port50_rows = [
    ["Kaohsiung", "TWKHH", "高雄港", "台灣 (Taiwan)", "遠東 / 亞洲樞紐"],
    ["Keelung", "TWKEL", "基隆港", "台灣 (Taiwan)", "亞洲近洋 / 遠東線"],
    ["Taichung", "TWTXG", "台中港", "台灣 (Taiwan)", "亞洲近洋 / 兩岸線"],
    ["Taipei", "TWTPE", "台北港", "台灣 (Taiwan)", "遠東幹線 / 近洋線"],
    ["Shanghai", "CNSHA", "上海港", "中國 (China)", "全球第一大貨櫃港"],
    ["Ningbo-Zhoushan", "CNNGB", "寧波舟山港", "中國 (China)", "遠東幹線航線樞紐"],
    ["Shenzhen", "CNSZX", "深圳港", "中國 (China)", "華南珠三角樞紐港"],
    ["Guangzhou", "CNGZG", "廣州港", "中國 (China)", "華南集裝箱主要港"],
    ["Qingdao", "CNTAO", "青島港", "中國 (China)", "華北環渤海主要港"],
    ["Hong Kong", "HKHKG", "香港港", "香港 (Hong Kong)", "華南國際轉運樞紐"],
    ["Singapore", "SGSIN", "新加坡港", "新加坡 (Singapore)", "全球最大轉運樞紐"],
    ["Port Klang", "MYPKG", "巴生港", "馬來西亞 (Malaysia)", "東南亞樞紐港"],
    ["Tanjung Pelepas", "MYTPP", "丹戎帕拉帕斯港", "馬來西亞 (Malaysia)", "馬六甲海峽轉運港"],
    ["Busan", "KRPUS", "釜山港", "韓國 (Korea)", "東北亞跨太平洋樞紐"],
    ["Tokyo", "JPTYO", "東京港", "日本 (Japan)", "日本關東主要貨櫃港"],
    ["Yokohama", "JPYOK", "橫濱港", "日本 (Japan)", "日本京濱主要外貿港"],
    ["Kobe", "JPUKB", "神戶港", "日本 (Japan)", "日本關西主要港口"],
    ["Osaka", "JPOSA", "大阪港", "日本 (Japan)", "日本關西內海港"],
    ["Nagoya", "JPNGO", "名古屋港", "日本 (Japan)", "日本中部汽車外貿港"],
    ["Hai Phong", "VNHPH", "海防港", "越南 (Vietnam)", "北越主要對外門戶"],
    ["Ho Chi Minh City", "VNSGN", "胡志明市港", "越南 (Vietnam)", "南越主要貨櫃樞紐"],
    ["Laem Chabang", "THLCH", "林查班港", "泰國 (Thailand)", "泰國最大深水貨櫃港"],
    ["Bangkok", "THBKK", "曼谷港", "泰國 (Thailand)", "泰國傳統河道港"],
    ["Jakarta", "IDJKT", "雅加達港", "印尼 (Indonesia)", "印尼爪哇門戶港"],
    ["Manila", "PHMNL", "馬尼拉港", "菲律賓 (Philippines)", "菲律賓首都主要港"],
    ["Los Angeles", "USLAX", "洛杉磯港", "美國 (USA)", "北美美西最大貨櫃港"],
    ["Long Beach", "USLGB", "長灘港", "美國 (USA)", "北美美西第二大貨櫃港"],
    ["Oakland", "USOAK", "奧克蘭港", "美國 (USA)", "北美美西灣區樞紐港"],
    ["Seattle", "USSEA", "西雅圖港", "美國 (USA)", "北美西北太平洋門戶"],
    ["Tacoma", "USTIW", "塔科馬港", "美國 (USA)", "北美西北海運聯盟"],
    ["New York / New Jersey", "USNYC", "紐約紐澤西港", "美國 (USA)", "北美美東第一大港"],
    ["Savannah", "USSAV", "薩凡納港", "美國 (USA)", "北美東南沿岸快速港"],
    ["Norfolk", "USORF", "諾福克港", "美國 (USA)", "北美維吉尼亞主要港"],
    ["Houston", "USHOU", "休士頓港", "美國 (USA)", "北美墨西哥灣主要港"],
    ["Vancouver", "CAVAN", "溫哥華港", "加拿大 (Canada)", "加拿大太平洋門戶港"],
    ["Prince Rupert", "CAPRR", "魯珀特王子港", "加拿大 (Canada)", "加拿大北美鐵路聯運"],
    ["Rotterdam", "NLRTM", "鹿特丹港", "荷蘭 (Netherlands)", "歐洲第一大門戶樞紐"],
    ["Antwerp", "BEANR", "安特衛普港", "比利時 (Belgium)", "歐洲第二大貨櫃港"],
    ["Hamburg", "DEHAM", "漢堡港", "德國 (Germany)", "歐洲易北河樞紐港"],
    ["Bremen / Bremerhaven", "DEBRV", "不萊梅哈芬港", "德國 (Germany)", "歐洲大型貨櫃場站"],
    ["Felixstowe", "GBFLX", "費利克斯托港", "英國 (UK)", "英國最大貨櫃港"],
    ["Southampton", "GBSOU", "南安普敦港", "英國 (UK)", "英國南部深水貨櫃港"],
    ["Le Havre", "FRLEH", "利哈佛港", "法國 (France)", "法國西北主要外貿港"],
    ["Fos-sur-Mer / Marseille", "FRFOS", "福斯港/馬賽港", "法國 (France)", "法國地中海門戶港"],
    ["Valencia", "ESVLC", "瓦倫西亞港", "西班牙 (Spain)", "西地中海主要樞紐"],
    ["Barcelona", "ESBCN", "巴塞隆納港", "西班牙 (Spain)", "加泰隆尼亞航運中心"],
    ["Algeciras", "ESALG", "阿爾赫西拉斯港", "西班牙 (Spain)", "直布羅陀海峽轉運港"],
    ["Genoa", "ITGOA", "熱那亞港", "義大利 (Italy)", "北義大利主要貨櫃港"],
    ["Piraeus", "GRPIR", "比雷埃夫斯港", "希臘 (Greece)", "東地中海核心轉運港"],
    ["Jebel Ali", "AEJEA", "傑貝阿里港", "阿聯酋 (UAE)", "中東最大國際樞紐港"]
]
add_table_data(doc, port50_headers, port50_rows, col_widths=[1.5, 1.0, 1.2, 1.2, 1.3], caption="表 附錄-1  全球 50 大主要貨櫃港口 UN/LOCODE 代碼速查對照表")

add_h2(doc, "附錄二：國際 ISO 6346 貨櫃尺寸型號代碼標準對照表")
add_p(doc, "表 附錄-2 整理了海運最常見之貨櫃尺寸、ISO 官方標準型號代碼與一般常見通稱對照：")

cntr_headers = ["常見業界通稱", "ISO 6346 尺寸型號碼", "長度 x 寬度 x 高度 (呎)", "內部容積 (約 CBM)", "標準最大總重 (Payload KGS)"]
cntr_rows = [
    ["20GP / 20'DC (20呎普通乾貨櫃)", "22G1", "20' x 8' x 8'6\"", "約 33 CBM", "約 21,800 KGS"],
    ["40GP / 40'DC (40呎普通平櫃)", "42G1", "40' x 8' x 8'6\"", "約 67 CBM", "約 26,600 KGS"],
    ["40HQ / 40'HC (40呎超高貨櫃)", "45G1", "40' x 8' x 9'6\"", "約 76 CBM", "約 26,500 KGS"],
    ["45HQ / 45'HC (45呎特長高櫃)", "L5G1", "45' x 8' x 9'6\"", "約 86 CBM", "約 27,800 KGS"],
    ["20RF (20呎低溫冷凍櫃)", "22R1", "20' x 8' x 8'6\"", "約 28 CBM", "約 21,400 KGS"],
    ["40RF / 40'RH (40呎超高冷凍櫃)", "45R1", "40' x 8' x 9'6\"", "約 67 CBM", "約 26,200 KGS"],
    ["20OT (20呎開頂櫃)", "22U1", "20' x 8' x 8'6\"", "約 32 CBM", "約 21,500 KGS"],
    ["40OT (40呎開頂櫃)", "42U1", "40' x 8' x 8'6\"", "約 65 CBM", "約 26,300 KGS"],
    ["20FR (20呎平板折疊櫃)", "22P1", "20' x 8' x 8'6\"", "無固定箱體 (重機設備)", "約 31,000 KGS"],
    ["40FR (40呎平板折疊櫃)", "42P1", "40' x 8' x 8'6\"", "無固定箱體 (超大貨物)", "約 40,000 KGS"]
]
add_table_data(doc, cntr_headers, cntr_rows, col_widths=[1.8, 1.1, 1.4, 1.1, 1.2], caption="表 附錄-2  國際標準貨櫃 ISO 6346 尺寸型號代碼標準對照表")

add_h2(doc, "附錄三：隨附 Excel 練習檔工作表索引說明")
add_p(doc, "為方便讀者實機演練，本實戰教材隨附提供兩份高度結構化的微軟 Excel 活頁簿（.xlsx 格式）：")

add_p(doc, "一、教材示範專用檔：《航運行政Excel資料前處理_示範資料檔_30筆.xlsx》")
add_p(doc, "● 工作表 1【01_原始待處理資料(Dirty)】：包含本手冊第三章所述之 30 筆示範髒資料，所有資料瑕疵完整保留。")
add_p(doc, "● 工作表 2【02_前處理完成資料(Cleaned)】：包含八大前處理步驟執行完畢後的標準成果（29筆唯一主檔），可作為操作驗收的對照標準。")

add_p(doc, "二、學員實戰練習檔：《航運行政Excel資料前處理_綜合練習題庫_300筆.xlsx》")
add_p(doc, "● 工作表 1【01_300筆航運情境練習題庫(Raw)】：完整收錄四大情境模組（模組A 100筆、模組B 70筆、模組C 70筆、模組D 60筆），供行政同仁逐一按部就班實戰演練。")
add_p(doc, "● 工作表 2【02_基礎代碼對照表(Lookup)】：提供國際主要港口之代碼、標準中文站名、區域與預設免費期對照表，供跨表查詢函數調用。")
add_p(doc, "● 工作表 3【03_解題思路與公式指引(Solutions)】：收錄所有題型之推薦解題技巧、核心函數語法及防呆注意事項。")

# Save complete manual
out_word_path = "c:\\proj\\w12\\航運行政人員_MS_Excel基礎資料前處理實戰手冊.docx"
doc.save(out_word_path)
print(f"Successfully generated and saved complete manual to: {out_word_path}")
