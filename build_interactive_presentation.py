import os
import base64
import json

fig_dir = "c:\\proj\\w12\\figures"
images_base64 = {}

fig_names = [
    "fig1_excel_interface.png",
    "fig2_data_preprocessing_flow.png",
    "fig3_clean_vs_dirty_data.png",
    "fig4_text_to_columns_wizard.png",
    "fig5_remove_duplicates_dialog.png",
    "fig6_flash_fill_demo.png",
    "fig7_data_validation_dialog.png",
    "fig8_core_functions_architecture.png",
    "fig9_demo30_before_after.png",
    "fig10_scenario_300_overview.png"
]

for fname in fig_names:
    fpath = os.path.join(fig_dir, fname)
    if os.path.exists(fpath):
        with open(fpath, "rb") as f:
            b64_str = base64.b64encode(f.read()).decode("utf-8")
            images_base64[fname] = f"data:image/png;base64,{b64_str}"
        print(f"Encoded {fname}")
    else:
        print(f"Warning: {fpath} not found!")

html_template = """<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>航運行政人員 MS Excel 基礎資料分析前處理實戰指南 - 互動簡報</title>
  <script src="https://www.gstatic.com/antigravity/web/dev/tailwindcss.min.js"></script>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;700;900&family=JetBrains+Mono:wght@400;600&display=swap');
    
    body {
      font-family: 'Noto Sans TC', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      user-select: none;
    }
    .code-font {
      font-family: 'JetBrains Mono', Consolas, Monaco, monospace;
    }
    
    /* Presentation Slide Transitions */
    .slide-container {
      perspective: 1200px;
    }
    .slide-page {
      display: none;
      opacity: 0;
      transform: scale(0.97) translateY(10px);
      transition: all 0.35s cubic-bezier(0.16, 1, 0.3, 1);
    }
    .slide-page.active {
      display: flex;
      opacity: 1;
      transform: scale(1) translateY(0);
    }
    
    /* Custom scrollbar for slide content */
    .slide-scroll::-webkit-scrollbar {
      width: 6px;
      height: 6px;
    }
    .slide-scroll::-webkit-scrollbar-track {
      background: rgba(255, 255, 255, 0.05);
      border-radius: 4px;
    }
    .slide-scroll::-webkit-scrollbar-thumb {
      background: rgba(255, 255, 255, 0.2);
      border-radius: 4px;
    }
    .slide-scroll::-webkit-scrollbar-thumb:hover {
      background: rgba(0, 168, 232, 0.5);
    }

    /* Pulse animation */
    @keyframes subtle-pulse {
      0%, 100% { opacity: 1; transform: scale(1); }
      50% { opacity: 0.85; transform: scale(0.98); }
    }
    .badge-pulse {
      animation: subtle-pulse 3s infinite ease-in-out;
    }
  </style>
</head>
<body class="bg-[#0B1528] text-slate-100 min-h-screen flex flex-col justify-between overflow-x-hidden">

  <!-- TOP HEADER / NAV BAR -->
  <header class="w-full bg-[#0F1D38]/90 backdrop-blur-md border-b border-slate-700/60 px-6 py-3 flex items-center justify-between z-30 shadow-md">
    <div class="flex items-center space-x-3">
      <div class="w-9 h-9 rounded-lg bg-gradient-to-tr from-cyan-600 to-blue-500 flex items-center justify-center shadow-lg shadow-cyan-500/20">
        <svg class="w-5 h-5 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.2" d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
        </svg>
      </div>
      <div>
        <h1 class="text-sm font-bold tracking-wide text-white flex items-center gap-2">
          航運行政人員 MS Excel 基礎資料分析前處理實戰指南
          <span class="text-xs font-semibold px-2 py-0.5 rounded-full bg-cyan-500/20 text-cyan-300 border border-cyan-500/30">PPT 投影片簡報模式</span>
        </h1>
        <p class="text-xs text-slate-400">國際航運行政數位轉型培訓教材系列 | 附 30 筆示範 + 300 筆實戰情境全解</p>
      </div>
    </div>

    <!-- Quick Actions -->
    <div class="flex items-center space-x-2">
      <button onclick="toggleGridOverview()" class="px-3 py-1.5 rounded-md bg-slate-800 hover:bg-slate-700 text-xs font-medium text-slate-200 border border-slate-600 transition flex items-center gap-1.5">
        <svg class="w-4 h-4 text-cyan-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 6a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2H6a2 2 0 01-2-2V6zM14 6a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2h-2a2 2 0 01-2-2V6zM4 16a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2H6a2 2 0 01-2-2v-2zM14 16a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2h-2a2 2 0 01-2-2v-2z"></path>
        </svg>
        投影片總覽 (O)
      </button>

      <button onclick="toggleNotes()" id="notes-btn" class="px-3 py-1.5 rounded-md bg-slate-800 hover:bg-slate-700 text-xs font-medium text-slate-200 border border-slate-600 transition flex items-center gap-1.5">
        <svg class="w-4 h-4 text-amber-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
        </svg>
        重點講義備忘 (N)
      </button>

      <button onclick="toggleFullscreen()" class="px-3 py-1.5 rounded-md bg-cyan-700 hover:bg-cyan-600 text-xs font-medium text-white transition flex items-center gap-1.5 shadow">
        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 8V4m0 0h4M4 4l5 5m11-1V4m0 0h-4m4 0l-5 5M4 16v4m0 0h4m-4 0l5-5m11 5l-5-5m5 5v-4m0 4h-4"></path>
        </svg>
        全螢幕 (F)
      </button>
    </div>
  </header>

  <!-- MAIN PRESENTATION STAGE -->
  <main class="flex-1 flex items-center justify-center p-4 sm:p-6 lg:p-8 relative">
    
    <!-- 16:9 Aspect Ratio Slide Stage Box -->
    <div id="slide-stage" class="w-full max-w-6xl aspect-[16/9] min-h-[580px] bg-gradient-to-br from-[#122344] via-[#0F1D38] to-[#0A1428] rounded-2xl border border-slate-700/80 shadow-2xl relative overflow-hidden flex flex-col justify-between slide-container">

      <!-- SLIDE TOP BAR: Category & Breadcrumb -->
      <div class="px-8 pt-5 pb-2 flex items-center justify-between border-b border-slate-800/80 text-xs text-slate-400">
        <div class="flex items-center space-x-2">
          <span class="inline-block w-2.5 h-2.5 rounded-full bg-cyan-400"></span>
          <span id="slide-category" class="font-bold tracking-wider uppercase text-cyan-300">模組總覽</span>
          <span class="text-slate-600">/</span>
          <span id="slide-section" class="text-slate-300">主題簡報</span>
        </div>
        <div class="flex items-center space-x-3">
          <span class="text-slate-400">航運行政 Excel 基礎資料分析前處理</span>
          <span id="slide-counter-top" class="px-2 py-0.5 rounded bg-slate-800 font-mono text-cyan-300 font-semibold">01 / 16</span>
        </div>
      </div>

      <!-- SLIDE PAGES WRAPPER -->
      <div id="slides-wrapper" class="flex-1 p-6 sm:p-8 flex flex-col justify-center overflow-y-auto slide-scroll relative">

        <!-- ======================================================== -->
        <!-- SLIDE 1: 封面投影片 -->
        <!-- ======================================================== -->
        <div class="slide-page active flex-col justify-center items-center text-center h-full space-y-6" data-category="手冊總覽" data-section="封面導言" data-notes="本投影片為整體課程總覽封面。請向學員強調本手冊專注於『資料前處理』，而非複雜的巨集或VBA。這是數據分析最關鍵的第一步。">
          <div class="inline-flex items-center gap-2 px-3.5 py-1.5 rounded-full bg-cyan-900/40 border border-cyan-500/40 text-cyan-300 text-xs font-semibold tracking-wider">
            <span class="w-2 h-2 rounded-full bg-cyan-400 animate-ping"></span>
            MARITIME LOGISTICS & SHIPPING DATA PREPROCESSING
          </div>
          <h2 class="text-3xl sm:text-5xl font-black text-transparent bg-clip-text bg-gradient-to-r from-white via-cyan-100 to-cyan-400 tracking-tight leading-tight">
            航運行政人員 MS Excel<br>基礎資料分析前處理實戰指南
          </h2>
          <p class="max-w-3xl text-sm sm:text-base text-slate-300 leading-relaxed font-light">
            從操作介面認識、常見髒資料特徵剖析、30筆示範案例全流程逐步圖解，<br class="hidden sm:inline">
            到 300 筆四大航運實務情境題庫全攻略（出口清單 / 貨櫃滯期費 / 貨代帳單 / 港口船期）
          </p>
          
          <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 max-w-3xl w-full pt-4">
            <div class="bg-slate-800/60 p-3 rounded-xl border border-slate-700/60 text-left">
              <span class="text-xs text-slate-400 block">適用對象</span>
              <span class="text-xs font-semibold text-white">航運/承攬/報關行政初學</span>
            </div>
            <div class="bg-slate-800/60 p-3 rounded-xl border border-slate-700/60 text-left">
              <span class="text-xs text-slate-400 block">核心工具</span>
              <span class="text-xs font-semibold text-white">MS Excel 365 / 2021 / 2024</span>
            </div>
            <div class="bg-slate-800/60 p-3 rounded-xl border border-slate-700/60 text-left">
              <span class="text-xs text-slate-400 block">實例演練</span>
              <span class="text-xs font-semibold text-cyan-300">30筆示範 + 300筆實戰</span>
            </div>
            <div class="bg-slate-800/60 p-3 rounded-xl border border-slate-700/60 text-left">
              <span class="text-xs text-slate-400 block">國際規範</span>
              <span class="text-xs font-semibold text-emerald-400">UN/LOCODE & ISO 6346</span>
            </div>
          </div>
        </div>

        <!-- ======================================================== -->
        <!-- SLIDE 2: 為什麼航運行政必須精通「資料前處理」？ -->
        <!-- ======================================================== -->
        <div class="slide-page flex-col justify-start h-full space-y-4" data-category="第一章 介面與導覽" data-section="1.1 核心理念" data-notes="強調 GIGO 原理。航運行政最常遇到系統匯出文字帶空格或格式混亂，導致算不出總重量、VLOOKUP 爆掉 #N/A。前處理佔據數據分析 80% 時間。">
          <div>
            <span class="text-xs font-bold text-cyan-400 tracking-wider">CH1 / 1.1 背景與痛點</span>
            <h3 class="text-2xl sm:text-3xl font-extrabold text-white mt-1">航運行政人員為什麼必須精通「資料前處理」？</h3>
            <p class="text-xs sm:text-sm text-slate-300">數據分析黃金鐵律：『垃圾進，垃圾出』(Garbage In, Garbage Out, GIGO)</p>
          </div>

          <div class="grid grid-cols-1 md:grid-cols-3 gap-4 pt-2">
            <div class="bg-rose-950/30 border border-rose-500/30 rounded-xl p-4 flex flex-col justify-between">
              <div>
                <div class="w-8 h-8 rounded-lg bg-rose-500/20 text-rose-400 flex items-center justify-center font-bold text-sm mb-3">1</div>
                <h4 class="text-sm font-bold text-white mb-2">提單號碼比對失敗 (#N/A)</h4>
                <p class="text-xs text-slate-300 leading-relaxed">
                  提單前後夾帶半形空格（如 <code class="bg-black/40 px-1 text-rose-300 font-mono">" EMCU20491823 "</code>），VLOOKUP/XLOOKUP 直接判定為不同字串，導致漏單、無法追蹤貨櫃動態。
                </p>
              </div>
              <span class="text-[11px] text-rose-400 font-semibold mt-3 block">衝擊：海運發單延誤、客戶客訴</span>
            </div>

            <div class="bg-amber-950/30 border border-amber-500/30 rounded-xl p-4 flex flex-col justify-between">
              <div>
                <div class="w-8 h-8 rounded-lg bg-amber-500/20 text-amber-400 flex items-center justify-center font-bold text-sm mb-3">2</div>
                <h4 class="text-sm font-bold text-white mb-2">重量與運費加總為 0</h4>
                <p class="text-xs text-slate-300 leading-relaxed">
                  貨物毛重輸入為 <code class="bg-black/40 px-1 text-amber-300 font-mono">"18,450.50 KGS"</code>，被 Excel 強制判定為「文字型態」，SUM 加總忽略不計顯示為 0，無法核算配載重心。
                </p>
              </div>
              <span class="text-[11px] text-amber-400 font-semibold mt-3 block">衝擊：運費少計、船舶配載失衡</span>
            </div>

            <div class="bg-blue-950/30 border border-blue-500/30 rounded-xl p-4 flex flex-col justify-between">
              <div>
                <div class="w-8 h-8 rounded-lg bg-blue-500/20 text-blue-400 flex items-center justify-center font-bold text-sm mb-3">3</div>
                <h4 class="text-sm font-bold text-white mb-2">滯期天數與航期算錯</h4>
                <p class="text-xs text-slate-300 leading-relaxed">
                  將開航日登記為 <code class="bg-black/40 px-1 text-blue-300 font-mono">"2026.11.05"</code> 文字，無法執行日期減法，導致貨櫃延滯費（D&D）免收期判定錯誤。
                </p>
              </div>
              <span class="text-[11px] text-blue-400 font-semibold mt-3 block">衝擊：高額滯期費爭議、帳單糾紛</span>
            </div>
          </div>

          <div class="bg-cyan-950/40 border border-cyan-500/40 rounded-xl p-4 flex items-center gap-4 mt-2">
            <div class="p-2.5 bg-cyan-500/20 text-cyan-300 rounded-lg shrink-0">
              <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>
            </div>
            <p class="text-xs sm:text-sm text-cyan-200">
              <span class="font-bold text-white">關鍵認知：</span>資料前處理佔據資料分析與統計報表高達 <strong class="text-cyan-300 font-bold underline">80% 的工作時間</strong>！唯有清洗出高質量的資料資產，後續的樞紐分析表（Pivot Table）與視覺化圖表才能得出正確結論。
            </p>
          </div>
        </div>

        <!-- ======================================================== -->
        <!-- SLIDE 3: Excel 視窗工作介面全景導覽 -->
        <!-- ======================================================== -->
        <div class="slide-page flex-col justify-start h-full space-y-4" data-category="第一章 介面與導覽" data-section="1.2 介面全景" data-notes="詳細介紹功能區 Ribbon、名稱方塊 Name Box（定位 B2）、資料編輯列 Formula Bar（檢視真實公式）、工作表分頁及狀態列（右鍵可查看即時加總與計數）。">
          <div class="flex items-center justify-between">
            <div>
              <span class="text-xs font-bold text-cyan-400 tracking-wider">CH1 / 1.2 視窗環境認識</span>
              <h3 class="text-2xl font-extrabold text-white">Excel 365 視窗工作介面全景導覽</h3>
            </div>
            <span class="text-xs px-2.5 py-1 bg-slate-800 text-slate-300 rounded border border-slate-700">圖 1-1 介面元素指引</span>
          </div>

          <div class="grid grid-cols-1 lg:grid-cols-12 gap-5 items-center">
            <!-- Left Image with modal trigger -->
            <div class="lg:col-span-7 bg-slate-900/60 p-2 rounded-xl border border-slate-700 flex flex-col items-center">
              <img src="IMAGE_fig1_excel_interface.png" alt="Excel 365 介面導覽" class="rounded-lg max-h-[300px] w-auto object-contain cursor-zoom-in shadow" onclick="zoomImage(this.src)">
              <span class="text-[11px] text-slate-400 mt-2">點擊上方圖表可放大檢視詳細標註</span>
            </div>

            <!-- Right Key Concepts -->
            <div class="lg:col-span-5 space-y-2.5 text-xs text-slate-200">
              <div class="bg-slate-800/80 p-3 rounded-lg border-l-4 border-cyan-400">
                <span class="font-bold text-cyan-300 block mb-0.5">1. 功能區 (Ribbon) 與「資料」標籤</span>
                <span>包含取得資料、排序篩選，以及前處理必備的【資料工具群】（資料剖析、移除重複項、閃電填入、資料驗證）。</span>
              </div>
              <div class="bg-slate-800/80 p-3 rounded-lg border-l-4 border-blue-400">
                <span class="font-bold text-blue-300 block mb-0.5">2. 名稱方塊 (Name Box) & 資料編輯列 (Formula Bar)</span>
                <span>名稱方塊顯示游標座標（如 B2）；資料編輯列忠實顯示儲存格底層的「真實公式」或未格式化純值。</span>
              </div>
              <div class="bg-slate-800/80 p-3 rounded-lg border-l-4 border-emerald-400">
                <span class="font-bold text-emerald-300 block mb-0.5">3. 工作表索引標籤 (Sheet Tabs)</span>
                <span>養成多表分離好習慣：將原始待處理檔（Dirty）與清洗完成主檔（Cleaned）分開存放。</span>
              </div>
              <div class="bg-slate-800/80 p-3 rounded-lg border-l-4 border-amber-400">
                <span class="font-bold text-amber-300 block mb-0.5">4. 底端狀態列 (Status Bar)</span>
                <span>選取欄位時自動即時顯示「筆數計數」、「加總 SUM」、「平均值」，是核對貨物總重的第一防線。</span>
              </div>
            </div>
          </div>
        </div>

        <!-- ======================================================== -->
        <!-- SLIDE 4: 航運文書處理必備快速鍵速查 -->
        <!-- ======================================================== -->
        <div class="slide-page flex-col justify-start h-full space-y-4" data-category="第一章 介面與導覽" data-section="1.4 實務快速鍵" data-notes="快速鍵是提升處理數千筆提單效率的神器。重點強調 Ctrl+E 閃電填入、Ctrl+Shift+L 快速篩選、F4 重複/絕對引用、以及 Ctrl+Alt+V 選擇性貼上值。">
          <div>
            <span class="text-xs font-bold text-cyan-400 tracking-wider">CH1 / 1.4 鍵盤操作心法</span>
            <h3 class="text-2xl font-extrabold text-white">航運文書處理必備快速鍵速查清單</h3>
            <p class="text-xs text-slate-300">拋棄繁瑣滑鼠切換，在數千筆貨櫃艙單中保持高速流暢節奏</p>
          </div>

          <div class="grid grid-cols-2 md:grid-cols-4 gap-3 pt-1">
            <div class="bg-slate-800/90 p-3.5 rounded-xl border border-slate-700 hover:border-cyan-500/50 transition">
              <span class="px-2 py-0.5 bg-cyan-900/60 border border-cyan-400/30 text-cyan-300 font-mono text-xs font-bold rounded inline-block mb-2">Ctrl + E</span>
              <h4 class="text-xs font-bold text-white">閃電填入 (Flash Fill)</h4>
              <p class="text-[11px] text-slate-400 mt-1 leading-snug">手打一個範例後秒速按 Ctrl+E，自動智慧切分全欄船名或航次！</p>
            </div>

            <div class="bg-slate-800/90 p-3.5 rounded-xl border border-slate-700 hover:border-cyan-500/50 transition">
              <span class="px-2 py-0.5 bg-blue-900/60 border border-blue-400/30 text-blue-300 font-mono text-xs font-bold rounded inline-block mb-2">Ctrl + Shift + L</span>
              <h4 class="text-xs font-bold text-white">快速啟用 / 關閉篩選</h4>
              <p class="text-[11px] text-slate-400 mt-1 leading-snug">一秒在標題列加上下拉按鈕，便於排查空值、重複值或異常提單。</p>
            </div>

            <div class="bg-slate-800/90 p-3.5 rounded-xl border border-slate-700 hover:border-cyan-500/50 transition">
              <span class="px-2 py-0.5 bg-emerald-900/60 border border-emerald-400/30 text-emerald-300 font-mono text-xs font-bold rounded inline-block mb-2">Ctrl + Alt + V</span>
              <h4 class="text-xs font-bold text-white">選擇性貼上 (Paste Special)</h4>
              <p class="text-[11px] text-slate-400 mt-1 leading-snug">清洗完公式後，貼上「值 (V)」切斷公式相依，轉換為純淨靜態主檔。</p>
            </div>

            <div class="bg-slate-800/90 p-3.5 rounded-xl border border-slate-700 hover:border-cyan-500/50 transition">
              <span class="px-2 py-0.5 bg-amber-900/60 border border-amber-400/30 text-amber-300 font-mono text-xs font-bold rounded inline-block mb-2">F4 / Ctrl + H</span>
              <h4 class="text-xs font-bold text-white">重複上一步 / 尋找取代</h4>
              <p class="text-[11px] text-slate-400 mt-1 leading-snug">F4 鎖定儲存格座標；Ctrl+H 批次把點號 '.' 或破折號全部替換掉。</p>
            </div>

            <div class="bg-slate-800/90 p-3.5 rounded-xl border border-slate-700 hover:border-cyan-500/50 transition">
              <span class="px-2 py-0.5 bg-slate-700 text-slate-200 font-mono text-xs font-bold rounded inline-block mb-2">Ctrl + A</span>
              <h4 class="text-xs font-bold text-white">全選當前連續資料表</h4>
              <p class="text-[11px] text-slate-400 mt-1 leading-snug">游標停在清單任一格，按 Ctrl+A 瞬間選取整張裝船總表。</p>
            </div>

            <div class="bg-slate-800/90 p-3.5 rounded-xl border border-slate-700 hover:border-cyan-500/50 transition">
              <span class="px-2 py-0.5 bg-slate-700 text-slate-200 font-mono text-xs font-bold rounded inline-block mb-2">Ctrl + 方向鍵</span>
              <h4 class="text-xs font-bold text-white">跳至資料極限端</h4>
              <p class="text-[11px] text-slate-400 mt-1 leading-snug">瞬間跳轉至第 1 列或最後一列（如第 5,000 列），無須滾動滑鼠。</p>
            </div>

            <div class="bg-slate-800/90 p-3.5 rounded-xl border border-slate-700 hover:border-cyan-500/50 transition">
              <span class="px-2 py-0.5 bg-slate-700 text-slate-200 font-mono text-xs font-bold rounded inline-block mb-2">F2 鍵</span>
              <h4 class="text-xs font-bold text-white">直接進入編輯模式</h4>
              <p class="text-[11px] text-slate-400 mt-1 leading-snug">游標直接跳入儲存格末端編輯文字，無須手動連點滑鼠兩下。</p>
            </div>

            <div class="bg-slate-800/90 p-3.5 rounded-xl border border-slate-700 hover:border-cyan-500/50 transition">
              <span class="px-2 py-0.5 bg-slate-700 text-slate-200 font-mono text-xs font-bold rounded inline-block mb-2">Alt + Enter</span>
              <h4 class="text-xs font-bold text-white">儲存格內強制換行</h4>
              <p class="text-[11px] text-slate-400 mt-1 leading-snug">在同一儲存格內將託運人名稱與地址分行排版，方便報關顯示。</p>
            </div>
          </div>
        </div>

        <!-- ======================================================== -->
        <!-- SLIDE 5: 什麼是乾淨資料 (Tidy Data)？航運數據三大原則 -->
        <!-- ======================================================== -->
        <div class="slide-page flex-col justify-start h-full space-y-4" data-category="第二章 核心常識" data-section="2.1 乾淨資料結構" data-notes="解釋 Tidy Data 三原則：每一列是一筆觀測值（提單）、每一欄是一個變數（起運港）、每個儲存格是單一純粹原子值。不可把船名與航次塞在同一格。">
          <div>
            <span class="text-xs font-bold text-cyan-400 tracking-wider">CH2 / 2.1 數據結構理論</span>
            <h3 class="text-2xl font-extrabold text-white">什麼是「乾淨資料」(Tidy Data)？航運數據三大基本原則</h3>
            <p class="text-xs text-slate-300">依據現代資料科學標準，建立高可用性航運資料庫的底層準則</p>
          </div>

          <div class="grid grid-cols-1 md:grid-cols-3 gap-5 pt-3">
            <div class="bg-gradient-to-b from-blue-900/30 to-slate-900/60 p-5 rounded-2xl border border-blue-500/30 flex flex-col justify-between">
              <div>
                <span class="text-2xl font-black text-cyan-400">01</span>
                <h4 class="text-base font-bold text-white mt-1 mb-2">原則一：每一列 (Row) 代表獨立實體</h4>
                <p class="text-xs text-slate-300 leading-relaxed">
                  在出口裝船總表中，每一列應且僅能代表一筆獨立提單（B/L No.）或一個實體貨櫃記錄。
                </p>
                <div class="mt-3 p-2 bg-black/40 rounded border border-slate-700 text-[11px] font-mono text-slate-300">
                  <span class="text-rose-400">[X]</span> 不可將2個櫃號合併寫同一列<br>
                  <span class="text-emerald-400">[√]</span> 每一列為單一櫃號獨立流水號
                </div>
              </div>
            </div>

            <div class="bg-gradient-to-b from-indigo-900/30 to-slate-900/60 p-5 rounded-2xl border border-indigo-500/30 flex flex-col justify-between">
              <div>
                <span class="text-2xl font-black text-indigo-400">02</span>
                <h4 class="text-base font-bold text-white mt-1 mb-2">原則二：每一欄 (Col) 代表單一業務屬性</h4>
                <p class="text-xs text-slate-300 leading-relaxed">
                  起運港 (POL)、目的港 (POD)、結關日 (Cutoff)、毛重 (KGS) 等必須各自獨立成直欄。
                </p>
                <div class="mt-3 p-2 bg-black/40 rounded border border-slate-700 text-[11px] font-mono text-slate-300">
                  <span class="text-rose-400">[X]</span> 欄名: '船名與航次 (合併)'<br>
                  <span class="text-emerald-400">[√]</span> 欄位A: 船名 | 欄位B: 航次
                </div>
              </div>
            </div>

            <div class="bg-gradient-to-b from-emerald-900/30 to-slate-900/60 p-5 rounded-2xl border border-emerald-500/30 flex flex-col justify-between">
              <div>
                <span class="text-2xl font-black text-emerald-400">03</span>
                <h4 class="text-base font-bold text-white mt-1 mb-2">原則三：每一個儲存格 (Cell) 僅含原子值</h4>
                <p class="text-xs text-slate-300 leading-relaxed">
                  數值必須是純數字，嚴禁夾雜文字單位（如 '15,000' 而不是 '15,000 KGS'）。
                </p>
                <div class="mt-3 p-2 bg-black/40 rounded border border-slate-700 text-[11px] font-mono text-slate-300">
                  <span class="text-rose-400">[X]</span> 儲存格內容: 'USD 2,450.00'<br>
                  <span class="text-emerald-400">[√]</span> 數值: 2450.00 | 幣別欄: USD
                </div>
              </div>
            </div>
          </div>
        </div>

        <!-- ======================================================== -->
        <!-- SLIDE 6: 儲存格資料型態原理與實務陷阱 -->
        <!-- ======================================================== -->
        <div class="slide-page flex-col justify-start h-full space-y-4" data-category="第二章 核心常識" data-section="2.2 資料型態" data-notes="詳細說明文字靠左、數值靠右。日期本質是自1900年起算的整數序列值，存成文字就無法相減。介紹 #N/A, #VALUE!, #REF! 錯誤代碼與處理方式。">
          <div>
            <span class="text-xs font-bold text-cyan-400 tracking-wider">CH2 / 2.2 底層運算原理</span>
            <h3 class="text-2xl font-extrabold text-white">儲存格資料型態原理與實務陷阱</h3>
            <p class="text-xs text-slate-300">了解文字、數值、日期與錯誤值的本質差異，避開運算失誤</p>
          </div>

          <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
            <!-- Left: Types -->
            <div class="bg-slate-900/80 p-4 rounded-xl border border-slate-700 space-y-2.5 text-xs">
              <h4 class="font-bold text-cyan-300 text-sm border-b border-slate-700 pb-1.5 flex items-center justify-between">
                <span>Excel 四大主要資料型態</span>
                <span class="text-[11px] font-normal text-slate-400">觀察預設對齊方式判斷</span>
              </h4>
              <div class="flex items-start gap-3 bg-slate-800/60 p-2 rounded">
                <span class="px-2 py-0.5 bg-blue-500/20 text-blue-300 font-bold rounded">文字 Text</span>
                <span class="text-slate-300">預設<strong>靠左對齊</strong>。純字串組合，不可直接做數學加減乘除。如提單號、託運人名稱、櫃號。</span>
              </div>
              <div class="flex items-start gap-3 bg-slate-800/60 p-2 rounded">
                <span class="px-2 py-0.5 bg-emerald-500/20 text-emerald-300 font-bold rounded">數值 Number</span>
                <span class="text-slate-300">預設<strong>靠右對齊</strong>。純浮點數或整數，可加總與樞紐分析。若夾雜英文單位將被強制降級為文字。</span>
              </div>
              <div class="flex items-start gap-3 bg-slate-800/60 p-2 rounded">
                <span class="px-2 py-0.5 bg-purple-500/20 text-purple-300 font-bold rounded">日期 Date</span>
                <span class="text-slate-300">預設<strong>靠右對齊</strong>。底層為自 1900/1/1 起算的連續序列整數，故「兩日期相減可直接計算天數」。</span>
              </div>
              <div class="flex items-start gap-3 bg-slate-800/60 p-2 rounded">
                <span class="px-2 py-0.5 bg-amber-500/20 text-amber-300 font-bold rounded">布林 Boolean</span>
                <span class="text-slate-300">預設<strong>置中對齊</strong>。僅 TRUE (真) 或 FALSE (偽) 兩種狀態，由邏輯判斷式生成。</span>
              </div>
            </div>

            <!-- Right: Errors -->
            <div class="bg-slate-900/80 p-4 rounded-xl border border-slate-700 space-y-2 text-xs">
              <h4 class="font-bold text-rose-300 text-sm border-b border-slate-700 pb-1.5 flex items-center justify-between">
                <span>航運高頻錯誤代碼 (#) 與解方</span>
                <span class="text-[11px] font-normal text-slate-400">掌握解法不再慌張</span>
              </h4>
              <div class="p-2 rounded bg-rose-950/30 border border-rose-500/20">
                <div class="flex justify-between font-mono font-bold text-rose-300">
                  <span>#N/A (查無資料)</span>
                  <span class="text-[11px] font-sans text-slate-400">VLOOKUP 比對失敗</span>
                </div>
                <p class="text-[11px] text-slate-300 mt-1">成因：提單號有隱藏空格。解法：使用 <code class="text-cyan-300 font-mono">TRIM()</code> 去空白或外層包 <code class="text-cyan-300 font-mono">IFERROR()</code>。</p>
              </div>
              <div class="p-2 rounded bg-amber-950/30 border border-amber-500/20">
                <div class="flex justify-between font-mono font-bold text-amber-300">
                  <span>#VALUE! (型態錯誤)</span>
                  <span class="text-[11px] font-sans text-slate-400">運算元不相容</span>
                </div>
                <p class="text-[11px] text-slate-300 mt-1">成因：拿包含文字的 '18450 KGS' 去做乘法。解法：用 <code class="text-cyan-300 font-mono">SUBSTITUTE()</code> 消除單位轉純數值。</p>
              </div>
              <div class="p-2 rounded bg-slate-800/60 border border-slate-700">
                <div class="flex justify-between font-mono font-bold text-slate-300">
                  <span>#REF! (參照損毀)</span>
                  <span class="text-[11px] font-sans text-slate-400">欄位已被誤刪除</span>
                </div>
                <p class="text-[11px] text-slate-300 mt-1">成因：公式引用的欄列被直接刪除。解法：Ctrl+Z 復原或重新指定參照座標。</p>
              </div>
            </div>
          </div>
        </div>

        <!-- ======================================================== -->
        <!-- SLIDE 7: 航運實務常見「七大髒資料症狀」與衝擊 -->
        <!-- ======================================================== -->
        <div class="slide-page flex-col justify-start h-full space-y-4" data-category="第二章 核心常識" data-section="2.3 髒資料症狀" data-notes="對照圖2-2說明七大症狀：多餘空白、大小寫混亂、未原子化、夾雜計量單位、非標準日期、非標準代碼與重複記錄。">
          <div class="flex items-center justify-between">
            <div>
              <span class="text-xs font-bold text-cyan-400 tracking-wider">CH2 / 2.3 實務痛點檢視</span>
              <h3 class="text-2xl font-extrabold text-white">航運實務常見「七大髒資料症狀」與標準對策</h3>
            </div>
            <span class="text-xs px-2.5 py-1 bg-slate-800 text-slate-300 rounded border border-slate-700">圖 2-2 症狀體系表</span>
          </div>

          <div class="grid grid-cols-1 lg:grid-cols-12 gap-5 items-center">
            <div class="lg:col-span-6 bg-slate-900/60 p-2 rounded-xl border border-slate-700 flex flex-col items-center">
              <img src="IMAGE_fig3_clean_vs_dirty_data.png" alt="七大髒資料對照圖" class="rounded-lg max-h-[300px] w-auto object-contain cursor-zoom-in shadow" onclick="zoomImage(this.src)">
              <span class="text-[11px] text-slate-400 mt-2">點擊上方圖表可放大檢視詳細對照</span>
            </div>

            <div class="lg:col-span-6 space-y-2 text-xs">
              <div class="flex items-center justify-between bg-slate-800/80 p-2.5 rounded-lg border-l-4 border-rose-500">
                <span class="font-bold text-rose-300">症狀 1：多餘隱藏空白</span>
                <span class="font-mono text-slate-300">"  EMCU20491823  " → TRIM()</span>
              </div>
              <div class="flex items-center justify-between bg-slate-800/80 p-2.5 rounded-lg border-l-4 border-amber-500">
                <span class="font-bold text-amber-300">症狀 2：大小寫混亂不一</span>
                <span class="font-mono text-slate-300">"twkhh" / "uslgb" → UPPER()</span>
              </div>
              <div class="flex items-center justify-between bg-slate-800/80 p-2.5 rounded-lg border-l-4 border-blue-500">
                <span class="font-bold text-blue-300">症狀 3：複合欄位未原子化</span>
                <span class="font-mono text-slate-300">"EVER GIVEN / 055W" → 資料剖析</span>
              </div>
              <div class="flex items-center justify-between bg-slate-800/80 p-2.5 rounded-lg border-l-4 border-emerald-500">
                <span class="font-bold text-emerald-300">症狀 4：數值夾雜計量單位</span>
                <span class="font-mono text-slate-300">"18,450 KGS" → SUBSTITUTE+VALUE</span>
              </div>
              <div class="flex items-center justify-between bg-slate-800/80 p-2.5 rounded-lg border-l-4 border-purple-500">
                <span class="font-bold text-purple-300">症狀 5：日期格式五花八門</span>
                <span class="font-mono text-slate-300">"2026.11.05" → DATE 統一西元日</span>
              </div>
              <div class="flex items-center justify-between bg-slate-800/80 p-2.5 rounded-lg border-l-4 border-cyan-500">
                <span class="font-bold text-cyan-300">症狀 6：非標準代碼與俗稱</span>
                <span class="font-mono text-slate-300">"40'HC", "40 DRY" → XLOOKUP/驗證</span>
              </div>
              <div class="flex items-center justify-between bg-slate-800/80 p-2.5 rounded-lg border-l-4 border-pink-500">
                <span class="font-bold text-pink-300">症狀 7：重複提單重複登錄</span>
                <span class="font-mono text-slate-300">重複登錄項次 → 移除重複項</span>
              </div>
            </div>
          </div>
        </div>

        <!-- ======================================================== -->
        <!-- SLIDE 8: 資料前處理標準五步驟循環 SOP -->
        <!-- ======================================================== -->
        <div class="slide-page flex-col justify-start h-full space-y-4" data-category="第二章 核心常識" data-section="2.4 前處理 SOP" data-notes="展示五步驟標準作業程序：備份檢視 -> 字串雜質清理 -> 欄位原子化 -> 型態規格化 -> 去重驗證。避免手忙腳亂。">
          <div class="flex items-center justify-between">
            <div>
              <span class="text-xs font-bold text-cyan-400 tracking-wider">CH2 / 2.4 標準作業流程</span>
              <h3 class="text-2xl font-extrabold text-white">資料前處理核心五步循環 SOP 流程</h3>
            </div>
            <span class="text-xs px-2.5 py-1 bg-slate-800 text-slate-300 rounded border border-slate-700">圖 2-1 流程架構</span>
          </div>

          <div class="bg-slate-900/60 p-3 rounded-xl border border-slate-700 flex justify-center">
            <img src="IMAGE_fig2_data_preprocessing_flow.png" alt="五步SOP流程圖" class="rounded-lg max-h-[200px] w-auto object-contain cursor-zoom-in" onclick="zoomImage(this.src)">
          </div>

          <div class="grid grid-cols-5 gap-2.5 pt-1 text-center">
            <div class="bg-slate-800/80 p-2.5 rounded-lg border border-slate-700">
              <span class="text-xs font-bold text-blue-400 block mb-1">步驟 1</span>
              <span class="text-xs font-bold text-white block">結構檢視與備份</span>
              <p class="text-[10px] text-slate-400 mt-1">複製原始工作表，檢查標題列唯一性與格式。</p>
            </div>
            <div class="bg-slate-800/80 p-2.5 rounded-lg border border-slate-700">
              <span class="text-xs font-bold text-emerald-400 block mb-1">步驟 2</span>
              <span class="text-xs font-bold text-white block">字串雜質清理</span>
              <p class="text-[10px] text-slate-400 mt-1">TRIM 清除首尾空白，UPPER 英文全大寫。</p>
            </div>
            <div class="bg-slate-800/80 p-2.5 rounded-lg border border-slate-700">
              <span class="text-xs font-bold text-amber-400 block mb-1">步驟 3</span>
              <span class="text-xs font-bold text-white block">欄位原子化分割</span>
              <p class="text-[10px] text-slate-400 mt-1">資料剖析切分船名航次，分離件數與單位。</p>
            </div>
            <div class="bg-slate-800/80 p-2.5 rounded-lg border border-slate-700">
              <span class="text-xs font-bold text-purple-400 block mb-1">步驟 4</span>
              <span class="text-xs font-bold text-white block">型態規格化轉換</span>
              <p class="text-[10px] text-slate-400 mt-1">統一西元日期格式，去除單位符號轉純數值。</p>
            </div>
            <div class="bg-slate-800/80 p-2.5 rounded-lg border border-slate-700">
              <span class="text-xs font-bold text-cyan-400 block mb-1">步驟 5</span>
              <span class="text-xs font-bold text-white block">去重比對與驗證</span>
              <p class="text-[10px] text-slate-400 mt-1">移除重複提單，比對港口主檔，貼上值結案。</p>
            </div>
          </div>
        </div>

        <!-- ======================================================== -->
        <!-- SLIDE 9: 30筆航運示範案例：原始結構與問題診斷 -->
        <!-- ======================================================== -->
        <div class="slide-page flex-col justify-start h-full space-y-4" data-category="第三章 示範案例" data-section="3.1 示範資料診斷" data-notes="展示本手冊隨附的30筆示範資料表。點出項次4與項次28完全重複，提單夾雜空格與小寫，重量帶有KGS字眼等實務毛病。">
          <div class="flex items-center justify-between">
            <div>
              <span class="text-xs font-bold text-cyan-400 tracking-wider">CH3 / 3.1 案例背景剖析</span>
              <h3 class="text-2xl font-extrabold text-white">30筆航運示範案例：原始資料結構與髒資料特徵</h3>
            </div>
            <span class="text-xs px-2.5 py-1 bg-slate-800 text-slate-300 rounded border border-slate-700">表 3-1 原始資料預覽</span>
          </div>

          <div class="bg-slate-900/90 rounded-xl border border-slate-700 overflow-x-auto text-[11px] slide-scroll">
            <table class="w-full text-left border-collapse">
              <thead>
                <tr class="bg-slate-800 text-slate-300 border-b border-slate-700 font-semibold">
                  <th class="p-2 text-center">項次</th>
                  <th class="p-2">提單號碼 (髒)</th>
                  <th class="p-2">船名航次</th>
                  <th class="p-2">起運港</th>
                  <th class="p-2">目的港</th>
                  <th class="p-2">櫃號</th>
                  <th class="p-2">櫃型</th>
                  <th class="p-2 text-right">毛重 (帶單位)</th>
                  <th class="p-2 text-center">結關日</th>
                  <th class="p-2 text-right">運費 (帶字串)</th>
                  <th class="p-2 text-rose-400 font-bold">診斷瑕疵備註</th>
                </tr>
              </thead>
              <tbody class="divide-y divide-slate-800/80 font-mono text-slate-200">
                <tr class="hover:bg-slate-800/40">
                  <td class="p-2 text-center font-bold">1</td>
                  <td class="p-2 text-rose-300">"  EMCU20491823 "</td>
                  <td class="p-2 text-amber-300">EVER GIVEN / 0124-055W</td>
                  <td class="p-2 text-rose-300">TWKHH </td>
                  <td class="p-2">USLAX</td>
                  <td class="p-2">EMCU8273641</td>
                  <td class="p-2">40'HC</td>
                  <td class="p-2 text-right text-rose-300">18,450.50 KGS</td>
                  <td class="p-2 text-center">2026/11/02</td>
                  <td class="p-2 text-right text-rose-300">USD 2,450.00</td>
                  <td class="p-2 text-rose-400 font-sans">前後空格、複合船名、重量文字</td>
                </tr>
                <tr class="hover:bg-slate-800/40">
                  <td class="p-2 text-center font-bold">2</td>
                  <td class="p-2">OOCLU918273645</td>
                  <td class="p-2 text-amber-300">OOCL ROTTERDAM / 088E</td>
                  <td class="p-2 text-rose-300">tw tpe</td>
                  <td class="p-2 text-rose-300">DEHAM </td>
                  <td class="p-2">OOLU1928374</td>
                  <td class="p-2">40HQ</td>
                  <td class="p-2 text-right">12400</td>
                  <td class="p-2 text-center text-amber-300">2026-11-03</td>
                  <td class="p-2 text-right">1,850</td>
                  <td class="p-2 text-rose-400 font-sans">港口小寫含空格、減號日期</td>
                </tr>
                <tr class="hover:bg-slate-800/40">
                  <td class="p-2 text-center font-bold">3</td>
                  <td class="p-2">MSCU992817263</td>
                  <td class="p-2 text-amber-300">MSC ISABELLA / 2311W</td>
                  <td class="p-2 text-rose-300">twkhh</td>
                  <td class="p-2 text-rose-300"> uslgB</td>
                  <td class="p-2">MSCU3847291</td>
                  <td class="p-2">20GP</td>
                  <td class="p-2 text-right text-rose-300">9,800.00 KGS</td>
                  <td class="p-2 text-center text-rose-300">2026.11.04</td>
                  <td class="p-2 text-right text-rose-300">$1,250.00</td>
                  <td class="p-2 text-rose-400 font-sans">點號日期、美元符號、港口混雜</td>
                </tr>
                <tr class="hover:bg-slate-800/40 bg-amber-950/20">
                  <td class="p-2 text-center font-bold text-amber-300">4</td>
                  <td class="p-2 font-bold text-amber-300">YMLU019283746</td>
                  <td class="p-2">YM WELLHEAD / 221E</td>
                  <td class="p-2">TWTXG</td>
                  <td class="p-2">JPTYO</td>
                  <td class="p-2">YMLU5566778</td>
                  <td class="p-2">20'GP</td>
                  <td class="p-2 text-right">5600 KGS</td>
                  <td class="p-2 text-center">2026/11/05</td>
                  <td class="p-2 text-right">USD 950.00</td>
                  <td class="p-2 text-amber-300 font-sans font-bold">★ 與項次 28 完全重複！</td>
                </tr>
                <tr class="hover:bg-slate-800/40 bg-amber-950/20">
                  <td class="p-2 text-center font-bold text-amber-300">28</td>
                  <td class="p-2 font-bold text-amber-300">YMLU019283746</td>
                  <td class="p-2">YM WELLHEAD / 221E</td>
                  <td class="p-2">TWTXG</td>
                  <td class="p-2">JPTYO</td>
                  <td class="p-2">YMLU5566778</td>
                  <td class="p-2">20'GP</td>
                  <td class="p-2 text-right">5600 KGS</td>
                  <td class="p-2 text-center">2026/11/05</td>
                  <td class="p-2 text-right">USD 950.00</td>
                  <td class="p-2 text-amber-300 font-sans font-bold">★ 重複登錄無效垃圾資料</td>
                </tr>
              </tbody>
            </table>
          </div>

          <p class="text-xs text-slate-300">
            完整 30 筆示範活頁簿保存在：<code class="text-cyan-300 bg-black/40 px-1 py-0.5 rounded font-mono">c:\proj\w12\航運行政Excel資料前處理_示範資料檔_30筆.xlsx</code>
          </p>
        </div>

        <!-- ======================================================== -->
        <!-- SLIDE 10: 30筆示範案例：清洗流程實戰八大步驟 -->
        <!-- ======================================================== -->
        <div class="slide-page flex-col justify-start h-full space-y-4" data-category="第三章 示範案例" data-section="3.3~3.10 清洗八步" data-notes="逐一步驟詳細講解：1.工作表備份；2.TRIM+UPPER提單；3.移除重複項；4.資料剖析拆船名航次；5.UN/LOCODE標準碼；6.SUBSTITUTE+VALUE去單位；7.DATE標準日期；8.金額轉純數值。">
          <div>
            <span class="text-xs font-bold text-cyan-400 tracking-wider">CH3 / 3.3~3.10 逐步實戰教學</span>
            <h3 class="text-2xl font-extrabold text-white">30筆示範案例：資料前處理八大實戰步驟指南</h3>
            <p class="text-xs text-slate-300">掌握關鍵功能與公式組合，逐一擊破航運資料瑕疵</p>
          </div>

          <div class="grid grid-cols-2 md:grid-cols-4 gap-3 text-xs">
            <div class="bg-slate-900/90 p-3 rounded-xl border border-slate-700">
              <span class="text-cyan-400 font-bold block mb-1">步驟 1：備份工作表</span>
              <p class="text-slate-300 text-[11px]">右鍵工作表標籤選【移動或複製】-> 勾選【建立複本】，命名為『清洗操作工作表』。</p>
            </div>
            <div class="bg-slate-900/90 p-3 rounded-xl border border-slate-700">
              <span class="text-cyan-400 font-bold block mb-1">步驟 2：提單去空格大寫</span>
              <div class="font-mono bg-black/40 p-1.5 rounded text-cyan-300 text-[11px] mb-1">=UPPER(TRIM(B2))</div>
              <p class="text-slate-300 text-[11px]">雙擊填滿控制點向下複製，最後「貼上值」切斷公式。</p>
            </div>
            <div class="bg-slate-900/90 p-3 rounded-xl border border-slate-700">
              <span class="text-cyan-400 font-bold block mb-1">步驟 3：移除重複提單</span>
              <p class="text-slate-300 text-[11px]">【資料】->【移除重複項】，僅勾選【提單號碼】，成功剔除項次 28，保留 29 筆唯一值。</p>
            </div>
            <div class="bg-slate-900/90 p-3 rounded-xl border border-slate-700">
              <span class="text-cyan-400 font-bold block mb-1">步驟 4：船名航次剖析</span>
              <p class="text-slate-300 text-[11px]">【資料剖析】精靈以斜線 '/' 為分隔符號，或手打一個範例後按快速鍵 <code class="text-amber-300 font-mono">Ctrl+E</code>！</p>
            </div>
            <div class="bg-slate-900/90 p-3 rounded-xl border border-slate-700">
              <span class="text-cyan-400 font-bold block mb-1">步驟 5：港口碼標準化</span>
              <div class="font-mono bg-black/40 p-1.5 rounded text-cyan-300 text-[10px] mb-1">=UPPER(TRIM(D2))</div>
              <p class="text-slate-300 text-[11px]">確保消除內部與首尾空格，符合聯合國 UN/LOCODE 5碼大寫。</p>
            </div>
            <div class="bg-slate-900/90 p-3 rounded-xl border border-slate-700">
              <span class="text-cyan-400 font-bold block mb-1">步驟 6：毛重去文字轉值</span>
              <div class="font-mono bg-black/40 p-1 rounded text-cyan-300 text-[10px] mb-1">=VALUE(SUBSTITUTE(SUBSTITUTE(I2,"KGS",""),",",""))</div>
              <p class="text-slate-300 text-[11px]">先清除 KGS 與逗點，再用 VALUE 轉純數值可加總。</p>
            </div>
            <div class="bg-slate-900/90 p-3 rounded-xl border border-slate-700">
              <span class="text-cyan-400 font-bold block mb-1">步驟 7：統一標準西元日</span>
              <p class="text-slate-300 text-[11px]">Ctrl+H 將點號 '.' 取代為斜線 '/'，儲存格格式設定為 YYYY-MM-DD。</p>
            </div>
            <div class="bg-slate-900/90 p-3 rounded-xl border border-slate-700">
              <span class="text-cyan-400 font-bold block mb-1">步驟 8：運費純數值化</span>
              <div class="font-mono bg-black/40 p-1 rounded text-cyan-300 text-[10px] mb-1">=VALUE(SUBSTITUTE(SUBSTITUTE(N2,"USD",""),"$",""))</div>
              <p class="text-slate-300 text-[11px]">消除外幣符號轉純浮點數，條款統一 PREPAID/COLLECT。</p>
            </div>
          </div>
        </div>

        <!-- ======================================================== -->
        <!-- SLIDE 11: 30筆示範案例：清洗前後成果對比展示 -->
        <!-- ======================================================== -->
        <div class="slide-page flex-col justify-start h-full space-y-4" data-category="第三章 示範案例" data-section="3.11 成效對比" data-notes="清洗前與清洗後具體成效對照。清洗前總重加總為0，清洗後總重 484,700 KGS、運費總額 USD 64,550，且去重成功，符合樞紐分析標準。">
          <div class="flex items-center justify-between">
            <div>
              <span class="text-xs font-bold text-cyan-400 tracking-wider">CH3 / 3.11 成果驗收核對</span>
              <h3 class="text-2xl font-extrabold text-white">30筆示範案例：清洗前 (Dirty) 與 清洗後 (Clean) 成效對照</h3>
            </div>
            <span class="text-xs px-2.5 py-1 bg-slate-800 text-slate-300 rounded border border-slate-700">圖 3-1 成效總結</span>
          </div>

          <div class="grid grid-cols-1 lg:grid-cols-12 gap-5 items-center">
            <div class="lg:col-span-7 bg-slate-900/60 p-2 rounded-xl border border-slate-700 flex flex-col items-center">
              <img src="IMAGE_fig9_demo30_before_after.png" alt="清洗前後對照" class="rounded-lg max-h-[300px] w-auto object-contain cursor-zoom-in shadow" onclick="zoomImage(this.src)">
              <span class="text-[11px] text-slate-400 mt-2">點擊上方圖表可放大檢視詳細成果</span>
            </div>

            <div class="lg:col-span-5 space-y-3 text-xs">
              <div class="bg-rose-950/30 p-3.5 rounded-xl border border-rose-500/30">
                <span class="font-bold text-rose-300 text-sm block mb-1.5">[X] 清洗前原始狀態</span>
                <ul class="space-y-1 text-slate-300 list-disc list-inside text-[11px]">
                  <li>提單夾帶空格，比對拋出 #N/A 錯誤</li>
                  <li>毛重欄位加總 SUM 結果為 <strong>0.00</strong></li>
                  <li>日期混雜無法計算天數與航期</li>
                  <li>重複登錄 1 筆無效紀錄，虛增貨量</li>
                </ul>
              </div>

              <div class="bg-emerald-950/30 p-3.5 rounded-xl border border-emerald-500/30">
                <span class="font-bold text-emerald-300 text-sm block mb-1.5">[√] 清洗後乾淨主檔</span>
                <ul class="space-y-1 text-slate-300 list-disc list-inside text-[11px]">
                  <li>提單純淨大寫，XLOOKUP 比對成功率 <strong>100%</strong></li>
                  <li>毛重加總正確得出總貨重，可分析載量</li>
                  <li>船名與航次原子化獨立兩欄</li>
                  <li>精確保留 29 筆唯一真實營收主檔</li>
                </ul>
              </div>
            </div>
          </div>
        </div>

        <!-- ======================================================== -->
        <!-- SLIDE 12: 前處理四大必備選單功能操作 -->
        <!-- ======================================================== -->
        <div class="slide-page flex-col justify-start h-full space-y-4" data-category="第四章 功能與函數" data-section="4.1 選單工具" data-notes="介紹四大選單：資料剖析精靈（圖4-1）、移除重複項對話框（圖4-2）、閃電填入Ctrl+E（圖4-3）、資料驗證下拉選單防呆（圖4-4）。免寫複雜公式。">
          <div>
            <span class="text-xs font-bold text-cyan-400 tracking-wider">CH4 / 4.1 功能區選單操作手冊</span>
            <h3 class="text-2xl font-extrabold text-white">前處理四大必備選單功能深入解析</h3>
            <p class="text-xs text-slate-300">善用 Excel 內建資料工具，無須撰寫複雜公式即可秒速完成高難度清洗</p>
          </div>

          <div class="grid grid-cols-2 md:grid-cols-4 gap-3 pt-2">
            <div class="bg-slate-900/90 p-3 rounded-xl border border-slate-700 flex flex-col justify-between">
              <div>
                <img src="IMAGE_fig4_text_to_columns_wizard.png" alt="資料剖析" class="rounded h-28 w-full object-cover mb-2 cursor-zoom-in" onclick="zoomImage(this.src)">
                <h4 class="text-xs font-bold text-white mb-1">1. 資料剖析 (Text to Cols)</h4>
                <p class="text-[11px] text-slate-300">指定斜線、空格或破折號為分隔符號，精準將複合欄位切分至相鄰欄。</p>
              </div>
              <span class="text-[10px] text-cyan-300 mt-2 block font-mono">功能位置：資料 -> 資料剖析</span>
            </div>

            <div class="bg-slate-900/90 p-3 rounded-xl border border-slate-700 flex flex-col justify-between">
              <div>
                <img src="IMAGE_fig5_remove_duplicates_dialog.png" alt="移除重複項" class="rounded h-28 w-full object-cover mb-2 cursor-zoom-in" onclick="zoomImage(this.src)">
                <h4 class="text-xs font-bold text-white mb-1">2. 移除重複項 (Duplicates)</h4>
                <p class="text-[11px] text-slate-300">防呆秘訣：絕不可全選所有欄位，僅勾選「提單號」或「櫃號」唯一主鍵。</p>
              </div>
              <span class="text-[10px] text-cyan-300 mt-2 block font-mono">功能位置：資料 -> 移除重複項</span>
            </div>

            <div class="bg-slate-900/90 p-3 rounded-xl border border-slate-700 flex flex-col justify-between">
              <div>
                <img src="IMAGE_fig6_flash_fill_demo.png" alt="閃電填入" class="rounded h-28 w-full object-cover mb-2 cursor-zoom-in" onclick="zoomImage(this.src)">
                <h4 class="text-xs font-bold text-white mb-1">3. 快速填入 (Flash Fill)</h4>
                <p class="text-[11px] text-slate-300">智慧模式學習：手動鍵入 1 個示範後按 Ctrl+E，自動完成全欄字串拆解。</p>
              </div>
              <span class="text-[10px] text-amber-300 mt-2 block font-mono">快速鍵：Ctrl + E</span>
            </div>

            <div class="bg-slate-900/90 p-3 rounded-xl border border-slate-700 flex flex-col justify-between">
              <div>
                <img src="IMAGE_fig7_data_validation_dialog.png" alt="資料驗證" class="rounded h-28 w-full object-cover mb-2 cursor-zoom-in" onclick="zoomImage(this.src)">
                <h4 class="text-xs font-bold text-white mb-1">4. 資料驗證 (Validation)</h4>
                <p class="text-[11px] text-slate-300">建立儲存格下拉選單（限定 20GP, 40GP, 40HQ），徹底防呆阻絕錯誤輸入。</p>
              </div>
              <span class="text-[10px] text-emerald-300 mt-2 block font-mono">功能位置：資料 -> 資料驗證</span>
            </div>
          </div>
        </div>

        <!-- ======================================================== -->
        <!-- SLIDE 13: 航運前處理「十大核心函數」體系導圖 -->
        <!-- ======================================================== -->
        <div class="slide-page flex-col justify-start h-full space-y-4" data-category="第四章 功能與函數" data-section="4.2 十大函數" data-notes="展示十大核心函數架構圖4-5：字串清洗（TRIM, CLEAN, UPPER）、字串拆解（LEFT, MID, RIGHT, SUBSTITUTE）、數值轉型（VALUE, TEXT）、日期比對（DATE, XLOOKUP, IFERROR）。">
          <div class="flex items-center justify-between">
            <div>
              <span class="text-xs font-bold text-cyan-400 tracking-wider">CH4 / 4.2 函數公式速查</span>
              <h3 class="text-2xl font-extrabold text-white">航運資料前處理「十大核心函數」體系架構導圖</h3>
            </div>
            <span class="text-xs px-2.5 py-1 bg-slate-800 text-slate-300 rounded border border-slate-700">圖 4-5 函數體系</span>
          </div>

          <div class="grid grid-cols-1 lg:grid-cols-12 gap-5 items-center">
            <div class="lg:col-span-7 bg-slate-900/60 p-2 rounded-xl border border-slate-700 flex flex-col items-center">
              <img src="IMAGE_fig8_core_functions_architecture.png" alt="十大核心函數體系" class="rounded-lg max-h-[300px] w-auto object-contain cursor-zoom-in shadow" onclick="zoomImage(this.src)">
              <span class="text-[11px] text-slate-400 mt-2">點擊上方圖表可放大檢視詳細架構</span>
            </div>

            <div class="lg:col-span-5 space-y-2 text-xs">
              <div class="bg-slate-800/80 p-2.5 rounded-lg border-l-4 border-emerald-500">
                <span class="font-bold text-emerald-300">字串清洗類 (Text Cleaning)</span>
                <p class="text-slate-300 font-mono text-[11px] mt-0.5"><strong class="text-white">TRIM</strong> (去首尾空格) | <strong class="text-white">CLEAN</strong> (除換行符) | <strong class="text-white">UPPER</strong> (轉大寫)</p>
              </div>

              <div class="bg-slate-800/80 p-2.5 rounded-lg border-l-4 border-blue-500">
                <span class="font-bold text-blue-300">字串截取替換 (Extraction)</span>
                <p class="text-slate-300 font-mono text-[11px] mt-0.5"><strong class="text-white">LEFT / MID / RIGHT</strong> (定點截取) | <strong class="text-white">SUBSTITUTE</strong> (精準替換)</p>
              </div>

              <div class="bg-slate-800/80 p-2.5 rounded-lg border-l-4 border-amber-500">
                <span class="font-bold text-amber-300">型態數值轉換 (Type Casting)</span>
                <p class="text-slate-300 font-mono text-[11px] mt-0.5"><strong class="text-white">VALUE</strong> (文字型數字轉真數值) | <strong class="text-white">TEXT</strong> (自訂格式化輸出)</p>
              </div>

              <div class="bg-slate-800/80 p-2.5 rounded-lg border-l-4 border-purple-500">
                <span class="font-bold text-purple-300">日期重構與查詢防呆 (Lookup & Date)</span>
                <p class="text-slate-300 font-mono text-[11px] mt-0.5"><strong class="text-white">DATE</strong> (重構西元日) | <strong class="text-white">XLOOKUP</strong> (跨表查找) | <strong class="text-white">IFERROR</strong> (錯誤攔截)</p>
              </div>
            </div>
          </div>
        </div>

        <!-- ======================================================== -->
        <!-- SLIDE 14: 300筆實務情境題庫四大業務模組 -->
        <!-- ======================================================== -->
        <div class="slide-page flex-col justify-start h-full space-y-4" data-category="第五章 300筆題庫" data-section="5.1~5.5 四大模組" data-notes="詳細介紹隨附的300筆練習題庫：模組A出口裝船清單100筆；模組B貨櫃延滯費70筆（MAX避免負數天數）；模組C貨代帳單70筆（多幣別折合TWD加稅）；模組D港口船期60筆（UN/LOCODE與航行天數）。">
          <div class="flex items-center justify-between">
            <div>
              <span class="text-xs font-bold text-cyan-400 tracking-wider">CH5 / 5.1~5.5 實戰題庫解析</span>
              <h3 class="text-2xl font-extrabold text-white">300筆航運情境練習題庫四大業務模組全覽</h3>
            </div>
            <span class="text-xs px-2.5 py-1 bg-slate-800 text-slate-300 rounded border border-slate-700">圖 5-1 題庫架構</span>
          </div>

          <div class="grid grid-cols-1 md:grid-cols-2 gap-3.5 pt-1 text-xs">
            <div class="bg-slate-900/90 p-3.5 rounded-xl border border-blue-500/30 flex flex-col justify-between">
              <div>
                <div class="flex justify-between items-center mb-1">
                  <span class="text-xs font-bold text-blue-400">模組 A (項次 1 - 100)</span>
                  <span class="text-[10px] px-2 py-0.5 rounded bg-blue-500/20 text-blue-300 font-semibold">100 筆訂艙清單</span>
                </div>
                <h4 class="text-sm font-bold text-white mb-1.5">海運出口訂艙與裝船清單清理</h4>
                <p class="text-[11px] text-slate-300 leading-relaxed">
                  聚焦提單去空格、貨櫃連字號去除、件數與單位分離、毛重純數值萃取、日期格式標準化。
                </p>
              </div>
              <div class="mt-2.5 p-2 bg-black/40 rounded border border-slate-700 text-[11px] font-mono text-cyan-300">
                核心公式：=VALUE(SUBSTITUTE(SUBSTITUTE(I2,"KGS",""),",",""))
              </div>
            </div>

            <div class="bg-slate-900/90 p-3.5 rounded-xl border border-amber-500/30 flex flex-col justify-between">
              <div>
                <div class="flex justify-between items-center mb-1">
                  <span class="text-xs font-bold text-amber-400">模組 B (項次 101 - 170)</span>
                  <span class="text-[10px] px-2 py-0.5 rounded bg-amber-500/20 text-amber-300 font-semibold">70 筆場站記錄</span>
                </div>
                <h4 class="text-sm font-bold text-white mb-1.5">貨櫃滯期延滯費 (D&D) 前處理</h4>
                <p class="text-[11px] text-slate-300 leading-relaxed">
                  卸船與還櫃日期相減、<code class="text-amber-300 font-mono">IF(ISBLANK())</code> 未還櫃防呆算至今日、<code class="text-amber-300 font-mono">MAX(0, 天數-免費期)</code> 杜絕負數天數。
                </p>
              </div>
              <div class="mt-2.5 p-2 bg-black/40 rounded border border-slate-700 text-[11px] font-mono text-amber-300">
                核心公式：=MAX(0, IF(ISBLANK(M2), TODAY()-L2, M2-L2) - 7)
              </div>
            </div>

            <div class="bg-slate-900/90 p-3.5 rounded-xl border border-emerald-500/30 flex flex-col justify-between">
              <div>
                <div class="flex justify-between items-center mb-1">
                  <span class="text-xs font-bold text-emerald-400">模組 C (項次 171 - 240)</span>
                  <span class="text-[10px] px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 font-semibold">70 筆費用帳單</span>
                </div>
                <h4 class="text-sm font-bold text-white mb-1.5">貨代應收應付帳單幣別匯率清理</h4>
                <p class="text-[11px] text-slate-300 leading-relaxed">
                  外幣代碼 (USD/EUR/JPY) 分離、原幣純數值萃取、依基準匯率 <code class="text-emerald-300 font-mono">ROUND()</code> 折合新台幣與 5% 營業稅精算。
                </p>
              </div>
              <div class="mt-2.5 p-2 bg-black/40 rounded border border-slate-700 text-[11px] font-mono text-emerald-300">
                核心公式：=ROUND(原幣純金額 * 折合匯率, 0)
              </div>
            </div>

            <div class="bg-slate-900/90 p-3.5 rounded-xl border border-purple-500/30 flex flex-col justify-between">
              <div>
                <div class="flex justify-between items-center mb-1">
                  <span class="text-xs font-bold text-purple-400">模組 D (項次 241 - 300)</span>
                  <span class="text-[10px] px-2 py-0.5 rounded bg-purple-500/20 text-purple-300 font-semibold">60 筆航線船期</span>
                </div>
                <h4 class="text-sm font-bold text-white mb-1.5">跨國港口與船期表標準化</h4>
                <p class="text-[11px] text-slate-300 leading-relaxed">
                  非標準港口別名透過 <code class="text-purple-300 font-mono">XLOOKUP()</code> 跨表比對聯合國 UN/LOCODE 主檔、計算 ETD-ETA 航程天數。
                </p>
              </div>
              <div class="mt-2.5 p-2 bg-black/40 rounded border border-slate-700 text-[11px] font-mono text-purple-300">
                核心公式：=XLOOKUP(G2, 港口表!$A:$A, 港口表!$B:$B, "查無代碼")
              </div>
            </div>
          </div>
        </div>

        <!-- ======================================================== -->
        <!-- SLIDE 15: 航運資料品質控管三階段 SOP 檢核表 -->
        <!-- ======================================================== -->
        <div class="slide-page flex-col justify-start h-full space-y-4" data-category="第六章 品質控管" data-section="6.1 SOP 檢核表" data-notes="強調資料清洗 SOP 檢核表：清洗前備份、標題唯一；清洗中主鍵去空白、原子化、數值轉型；清洗後去重確認、貼上值切斷公式、業務邏輯審查。">
          <div>
            <span class="text-xs font-bold text-cyan-400 tracking-wider">CH6 / 6.1 品質保證指南</span>
            <h3 class="text-2xl font-extrabold text-white">航運行政人員資料前處理三階段 SOP 檢核表</h3>
            <p class="text-xs text-slate-300">遵循標準查核檢驗程序，杜絕資料遺失與計算偏差</p>
          </div>

          <div class="grid grid-cols-1 md:grid-cols-3 gap-4 pt-1">
            <div class="bg-slate-900/80 p-4 rounded-xl border-t-4 border-blue-500 text-xs flex flex-col justify-between">
              <div>
                <span class="text-xs font-bold text-blue-400 uppercase tracking-wide">第一階段</span>
                <h4 class="text-sm font-bold text-white mt-1 mb-2">清洗前 (Pre-check)</h4>
                <ul class="space-y-2 text-slate-300">
                  <li class="flex items-start gap-1.5">
                    <span class="text-blue-400 font-bold">[1]</span>
                    <span><strong>原始檔案備份</strong>：另存新檔並建立複製工作表，原始底稿永不變更。</span>
                  </li>
                  <li class="flex items-start gap-1.5">
                    <span class="text-blue-400 font-bold">[2]</span>
                    <span><strong>標題列唯一性檢查</strong>：確認第 1 列無合併儲存格、無重複欄名。</span>
                  </li>
                  <li class="flex items-start gap-1.5">
                    <span class="text-blue-400 font-bold">[3]</span>
                    <span><strong>初始總筆數記錄</strong>：透過狀態列記錄原始列數（如 300 列）。</span>
                  </li>
                </ul>
              </div>
              <span class="text-[11px] text-blue-400 mt-3 font-semibold">目標：確保原始資料底稿安全</span>
            </div>

            <div class="bg-slate-900/80 p-4 rounded-xl border-t-4 border-emerald-500 text-xs flex flex-col justify-between">
              <div>
                <span class="text-xs font-bold text-emerald-400 uppercase tracking-wide">第二階段</span>
                <h4 class="text-sm font-bold text-white mt-1 mb-2">清洗中 (In-process)</h4>
                <ul class="space-y-2 text-slate-300">
                  <li class="flex items-start gap-1.5">
                    <span class="text-emerald-400 font-bold">[4]</span>
                    <span><strong>主鍵去空白大寫</strong>：提單號、櫃號套用 TRIM 與 UPPER。</span>
                  </li>
                  <li class="flex items-start gap-1.5">
                    <span class="text-emerald-400 font-bold">[5]</span>
                    <span><strong>欄位原子化分割</strong>：船名航次、件數單位必須各自獨立一欄。</span>
                  </li>
                  <li class="flex items-start gap-1.5">
                    <span class="text-emerald-400 font-bold">[6]</span>
                    <span><strong>數值與日期真偽校驗</strong>：以 ISNUMBER 檢查重量與金額是否為真數值。</span>
                  </li>
                </ul>
              </div>
              <span class="text-[11px] text-emerald-400 mt-3 font-semibold">目標：達成原子化與型態純淨</span>
            </div>

            <div class="bg-slate-900/80 p-4 rounded-xl border-t-4 border-purple-500 text-xs flex flex-col justify-between">
              <div>
                <span class="text-xs font-bold text-purple-400 uppercase tracking-wide">第三階段</span>
                <h4 class="text-sm font-bold text-white mt-1 mb-2">清洗後 (Post-check)</h4>
                <ul class="space-y-2 text-slate-300">
                  <li class="flex items-start gap-1.5">
                    <span class="text-purple-400 font-bold">[7]</span>
                    <span><strong>重複記錄剔除確認</strong>：依主鍵移除重複，核對剔除筆數原因。</span>
                  </li>
                  <li class="flex items-start gap-1.5">
                    <span class="text-purple-400 font-bold">[8]</span>
                    <span><strong>公式轉純值 (貼上值)</strong>：複製 -> 選擇性貼上為『值』，解除運算負擔。</span>
                  </li>
                  <li class="flex items-start gap-1.5">
                    <span class="text-purple-400 font-bold">[9]</span>
                    <span><strong>業務邏輯審查</strong>：篩選排查毛重是否負數、開航日是否早於結關日。</span>
                  </li>
                </ul>
              </div>
              <span class="text-[11px] text-purple-400 mt-3 font-semibold">目標：確保分析資料合規就緒</span>
            </div>
          </div>
        </div>

        <!-- ======================================================== -->
        <!-- SLIDE 16: 總結、隨附資源與國際參考文獻 -->
        <!-- ======================================================== -->
        <div class="slide-page flex-col justify-start h-full space-y-4" data-category="手冊總結" data-section="6.2 參考文獻與資源" data-notes="總結全文：推薦學員開啟隨附的兩份 Excel 活頁簿進行實機操作練習。註明聯合國 UN/LOCODE、ISO 6346、ISO 8601、微軟官方文檔與 Hadley Wickham 乾淨資料論文等引用來源。">
          <div>
            <span class="text-xs font-bold text-cyan-400 tracking-wider">手冊總結 / 資源與規範</span>
            <h3 class="text-2xl font-extrabold text-white">總結：隨附學習資源、活頁簿清單與國際參考文獻</h3>
            <p class="text-xs text-slate-300">實機演練、持續深化數據分析技能，邁向航運數位化領航者</p>
          </div>

          <div class="grid grid-cols-1 md:grid-cols-2 gap-4 pt-1">
            <!-- Left: Files -->
            <div class="bg-slate-900/90 p-4 rounded-xl border border-cyan-500/30 space-y-2.5 text-xs">
              <h4 class="font-bold text-cyan-300 text-sm border-b border-slate-700 pb-1 flex items-center justify-between">
                <span>隨附正式活頁簿資源</span>
                <span class="text-[11px] font-normal text-slate-400">已存於本機工作區</span>
              </h4>
              <div class="p-2.5 bg-slate-800/80 rounded border border-slate-700">
                <span class="font-bold text-white block">1. 《航運行政Excel資料前處理_示範資料檔_30筆.xlsx》</span>
                <span class="text-[11px] text-slate-300 block mt-0.5">含【01_原始待處理(Dirty)】與【02_前處理完成(Cleaned)】對照表。</span>
              </div>
              <div class="p-2.5 bg-slate-800/80 rounded border border-slate-700">
                <span class="font-bold text-white block">2. 《航運行政Excel資料前處理_綜合練習題庫_300筆.xlsx》</span>
                <span class="text-[11px] text-slate-300 block mt-0.5">含 300 筆四大情境題庫、全球港口主檔對照表及全套函數解答指引。</span>
              </div>
              <div class="p-2.5 bg-slate-800/80 rounded border border-slate-700">
                <span class="font-bold text-white block">3. 《航運行政人員_MS_Excel基礎資料前處理實戰手冊.docx》</span>
                <span class="text-[11px] text-slate-300 block mt-0.5">完整 60 頁詳細圖文教學手冊（含封面、目錄、圖表與頁碼編排）。</span>
              </div>
            </div>

            <!-- Right: References -->
            <div class="bg-slate-900/90 p-4 rounded-xl border border-slate-700 space-y-2 text-xs">
              <h4 class="font-bold text-emerald-300 text-sm border-b border-slate-700 pb-1">國際權威參考文獻與引用來源</h4>
              <ul class="space-y-1.5 text-slate-300 text-[11px]">
                <li><strong class="text-white">● Microsoft Learn Documentation (2024)</strong>: Excel Functions by Category (Text, Lookup, Date/Time).</li>
                <li><strong class="text-white">● UN/CEFACT (UNECE)</strong>: UN/LOCODE: Codes for Trade and Transport Locations (Recommendation No. 16).</li>
                <li><strong class="text-white">● ISO 8601:2019</strong>: Data elements and interchange formats — Representation of dates and times.</li>
                <li><strong class="text-white">● ISO 6346:2022</strong>: Freight containers — Coding, identification and marking.</li>
                <li><strong class="text-white">● Hadley Wickham (2014)</strong>: Tidy Data. Journal of Statistical Software, 59(10).</li>
                <li><strong class="text-white">● FIATA (2023)</strong>: Multimodal Transport Guidelines and Standard Documentation.</li>
              </ul>
            </div>
          </div>

          <div class="p-3 bg-gradient-to-r from-cyan-950/60 to-blue-950/60 border border-cyan-500/30 rounded-xl text-center">
            <p class="text-xs text-slate-200">
              恭喜您完成《航運行政人員 MS Excel 基礎資料分析前處理實戰指南》投影片研讀！<br>
              請即刻開啟隨附的 Excel 活頁簿動手演練，體會乾淨資料帶來的高效與成就感！
            </p>
          </div>
        </div>

      </div>

      <!-- SLIDE BOTTOM CONTROL BAR -->
      <div class="px-6 py-3 bg-[#0D182E]/95 border-t border-slate-800/90 flex items-center justify-between z-20">
        <!-- Prev / Next Buttons -->
        <div class="flex items-center space-x-2">
          <button onclick="prevSlide()" id="btn-prev" class="px-3.5 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-xs font-semibold text-slate-200 border border-slate-700 transition flex items-center gap-1">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 19l-7-7 7-7"></path></svg>
            上一頁 (←)
          </button>
          <button onclick="nextSlide()" id="btn-next" class="px-4 py-1.5 rounded-lg bg-gradient-to-r from-cyan-600 to-blue-600 hover:from-cyan-500 hover:to-blue-500 text-xs font-semibold text-white shadow transition flex items-center gap-1">
            下一頁 (→)
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5l7 7-7 7"></path></svg>
          </button>
        </div>

        <!-- Progress Bar & Indicator -->
        <div class="flex-1 max-w-md mx-6 flex items-center space-x-3">
          <div class="w-full bg-slate-800 h-2 rounded-full overflow-hidden border border-slate-700/60">
            <div id="progress-bar" class="bg-gradient-to-r from-cyan-500 to-blue-500 h-full w-[6.25%] transition-all duration-300"></div>
          </div>
          <span id="slide-indicator" class="text-xs font-mono font-bold text-cyan-300 whitespace-nowrap">01 / 16</span>
        </div>

        <!-- Interactive Playground Button -->
        <div class="flex items-center space-x-2">
          <button onclick="toggleInteractiveLab()" class="px-3 py-1.5 rounded-lg bg-emerald-950 hover:bg-emerald-900 border border-emerald-500/40 text-emerald-300 text-xs font-semibold transition flex items-center gap-1.5">
            <span class="w-2 h-2 rounded-full bg-emerald-400 animate-ping"></span>
            🧪 互動清洗實驗室
          </button>
        </div>
      </div>

    </div>
  </main>

  <!-- ======================================================== -->
  <!-- MODAL: INTERACTIVE DATA PREPROCESSING LAB (即時清洗模擬器) -->
  <!-- ======================================================== -->
  <div id="lab-modal" class="fixed inset-0 bg-black/80 backdrop-blur-sm z-50 hidden flex items-center justify-center p-4">
    <div class="bg-slate-900 border border-slate-700 rounded-2xl max-w-2xl w-full p-6 shadow-2xl relative">
      <div class="flex items-center justify-between pb-3 border-b border-slate-800">
        <div class="flex items-center gap-2">
          <div class="w-8 h-8 rounded-lg bg-emerald-500/20 text-emerald-400 flex items-center justify-center font-bold text-base">🧪</div>
          <div>
            <h3 class="text-sm font-bold text-white">航運資料前處理「即時公式模擬器」</h3>
            <p class="text-xs text-slate-400">親自體驗 Excel 前處理函數如何瞬間化腐朽為神奇</p>
          </div>
        </div>
        <button onclick="toggleInteractiveLab()" class="text-slate-400 hover:text-white text-lg">✕</button>
      </div>

      <div class="space-y-4 py-4 text-xs">
        <div>
          <label class="block text-slate-300 font-semibold mb-1">1. 輸入或選取待清洗的髒資料字串：</label>
          <div class="flex gap-2">
            <input type="text" id="lab-input" value="  emcu20491823  " class="flex-1 bg-black/60 border border-slate-700 rounded px-3 py-2 text-white font-mono focus:border-cyan-500 outline-none">
            <button onclick="loadSample('blank')" class="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 rounded text-slate-300 border border-slate-700">空格範例</button>
            <button onclick="loadSample('vessel')" class="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 rounded text-slate-300 border border-slate-700">船名範例</button>
            <button onclick="loadSample('weight')" class="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 rounded text-slate-300 border border-slate-700">重量範例</button>
          </div>
        </div>

        <div>
          <label class="block text-slate-300 font-semibold mb-1.5">2. 點選欲測試套用的 Excel 清洗運算：</label>
          <div class="grid grid-cols-3 gap-2">
            <button onclick="applyFormula('trim_upper')" class="p-2 bg-slate-800 hover:bg-cyan-900/40 hover:border-cyan-500 border border-slate-700 rounded text-left transition">
              <span class="font-bold text-cyan-300 block">UPPER(TRIM())</span>
              <span class="text-[10px] text-slate-400">消除首尾空白並轉全大寫</span>
            </button>
            <button onclick="applyFormula('split_vessel')" class="p-2 bg-slate-800 hover:bg-cyan-900/40 hover:border-cyan-500 border border-slate-700 rounded text-left transition">
              <span class="font-bold text-amber-300 block">LEFT/FIND 剖析</span>
              <span class="text-[10px] text-slate-400">自斜線分離船名與航次</span>
            </button>
            <button onclick="applyFormula('clean_weight')" class="p-2 bg-slate-800 hover:bg-cyan-900/40 hover:border-cyan-500 border border-slate-700 rounded text-left transition">
              <span class="font-bold text-emerald-300 block">SUBSTITUTE+VALUE</span>
              <span class="text-[10px] text-slate-400">清除 KGS 與逗點轉純數值</span>
            </button>
          </div>
        </div>

        <div class="bg-black/60 p-4 rounded-xl border border-slate-800 space-y-2">
          <div class="flex justify-between items-center text-slate-400">
            <span>使用 Excel 等效公式：</span>
            <span id="lab-formula-name" class="font-mono text-cyan-400 font-bold">=UPPER(TRIM(A2))</span>
          </div>
          <div class="border-t border-slate-800 pt-2">
            <span class="text-slate-400 block mb-1">清洗後標準成果輸出：</span>
            <div id="lab-output" class="text-base font-mono font-bold text-emerald-400 bg-slate-900/80 p-2.5 rounded border border-emerald-500/30">
              EMCU20491823
            </div>
          </div>
        </div>
      </div>

      <div class="flex justify-end pt-2">
        <button onclick="toggleInteractiveLab()" class="px-4 py-1.5 bg-slate-800 hover:bg-slate-700 text-xs font-semibold text-slate-200 rounded">關閉視窗</button>
      </div>
    </div>
  </div>

  <!-- ======================================================== -->
  <!-- MODAL: SLIDE GRID OVERVIEW (縮圖總覽) -->
  <!-- ======================================================== -->
  <div id="grid-modal" class="fixed inset-0 bg-black/85 backdrop-blur-md z-50 hidden flex flex-col p-6 overflow-y-auto slide-scroll">
    <div class="max-w-6xl w-full mx-auto flex items-center justify-between pb-4 border-b border-slate-800">
      <div>
        <h3 class="text-lg font-bold text-white">投影片快速切換總覽 (Slide Navigator)</h3>
        <p class="text-xs text-slate-400">點擊任意卡片可直接跳轉至該投影片，按 ESC 或點擊關閉返回簡報</p>
      </div>
      <button onclick="toggleGridOverview()" class="px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-xs font-semibold text-slate-200 border border-slate-700">關閉總覽 (ESC)</button>
    </div>

    <div id="grid-cards-container" class="max-w-6xl w-full mx-auto grid grid-cols-2 sm:grid-cols-4 gap-4 py-6">
      <!-- Generated via JavaScript -->
    </div>
  </div>

  <!-- ======================================================== -->
  <!-- MODAL: IMAGE ZOOM LIGHTBOX -->
  <!-- ======================================================== -->
  <div id="image-modal" class="fixed inset-0 bg-black/90 backdrop-blur-md z-50 hidden flex flex-col items-center justify-center p-6" onclick="closeImageZoom()">
    <img id="modal-img" src="" alt="放大圖表" class="max-w-5xl max-h-[85vh] object-contain rounded-xl shadow-2xl border border-slate-700">
    <span class="text-xs text-slate-400 mt-3">點擊螢幕任意處關閉放大視窗</span>
  </div>

  <!-- ======================================================== -->
  <!-- SPEAKER NOTES FLOATING DRAWER -->
  <!-- ======================================================== -->
  <div id="speaker-notes" class="fixed bottom-16 right-8 max-w-md w-full bg-slate-900/95 border border-amber-500/40 rounded-xl p-4 shadow-2xl z-40 hidden backdrop-blur">
    <div class="flex items-center justify-between pb-2 border-b border-slate-800">
      <span class="text-xs font-bold text-amber-400 flex items-center gap-1.5">
        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path></svg>
        講師演說重點與實務叮嚀 (Speaker Notes)
      </span>
      <button onclick="toggleNotes()" class="text-slate-400 hover:text-white text-xs">✕</button>
    </div>
    <div id="notes-content" class="text-xs text-slate-300 mt-2.5 leading-relaxed max-h-40 overflow-y-auto slide-scroll">
      備忘錄內容載入中...
    </div>
  </div>

  <!-- JAVASCRIPT LOGIC -->
  <script>
    let currentSlide = 0;
    const slides = document.querySelectorAll('.slide-page');
    const totalSlides = slides.length;

    const titles = [
      "封面導言：航運行政前處理實戰指南",
      "CH1 / 為什麼必須精通「資料前處理」？",
      "CH1 / Excel 365 視窗工作介面全景導覽",
      "CH1 / 航運文書處理必備快速鍵速查",
      "CH2 / 什麼是乾淨資料 (Tidy Data)？三大原則",
      "CH2 / 儲存格資料型態原理與實務陷阱",
      "CH2 / 航運實務常見「七大髒資料症狀」",
      "CH2 / 資料前處理核心五步循環 SOP",
      "CH3 / 30筆航運示範案例原始結構與診斷",
      "CH3 / 示範案例：清洗流程實戰八大步驟",
      "CH3 / 示範案例：清洗前後成果對比展示",
      "CH4 / 前處理四大必備選單功能深入解析",
      "CH4 / 航運前處理「十大核心函數」體系架構",
      "CH5 / 300筆航運實戰題庫四大業務模組全覽",
      "CH6 / 航運資料品質控管三階段 SOP 檢核表",
      "總結：隨附活頁簿指引與國際參考文獻"
    ];

    function updateSlide() {
      slides.forEach((s, idx) => {
        if (idx === currentSlide) {
          s.classList.add('active');
        } else {
          s.classList.remove('active');
        }
      });

      // Update counters & headers
      const currentNum = String(currentSlide + 1).padStart(2, '0');
      const totalNum = String(totalSlides).padStart(2, '0');
      document.getElementById('slide-counter-top').textContent = `${currentNum} / ${totalNum}`;
      document.getElementById('slide-indicator').textContent = `${currentNum} / ${totalNum}`;

      // Update progress bar
      const pct = ((currentSlide + 1) / totalSlides) * 100;
      document.getElementById('progress-bar').style.width = pct + '%';

      // Update Category & Section
      const curElem = slides[currentSlide];
      document.getElementById('slide-category').textContent = curElem.dataset.category || '模組導覽';
      document.getElementById('slide-section').textContent = curElem.dataset.section || '章節內容';

      // Update Speaker Notes
      document.getElementById('notes-content').textContent = curElem.dataset.notes || '本頁尚無特殊講師備忘。';

      // Buttons state
      document.getElementById('btn-prev').disabled = (currentSlide === 0);
      document.getElementById('btn-prev').style.opacity = (currentSlide === 0) ? '0.4' : '1';
      document.getElementById('btn-next').disabled = (currentSlide === totalSlides - 1);
      document.getElementById('btn-next').style.opacity = (currentSlide === totalSlides - 1) ? '0.4' : '1';
    }

    function prevSlide() {
      if (currentSlide > 0) {
        currentSlide--;
        updateSlide();
      }
    }

    function nextSlide() {
      if (currentSlide < totalSlides - 1) {
        currentSlide++;
        updateSlide();
      }
    }

    function jumpToSlide(idx) {
      if (idx >= 0 && idx < totalSlides) {
        currentSlide = idx;
        updateSlide();
        document.getElementById('grid-modal').classList.add('hidden');
      }
    }

    // Keyboard Shortcuts
    window.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') {
        nextSlide();
      } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
        prevSlide();
      } else if (e.key === 'Home') {
        jumpToSlide(0);
      } else if (e.key === 'End') {
        jumpToSlide(totalSlides - 1);
      } else if (e.key === 'f' || e.key === 'F') {
        toggleFullscreen();
      } else if (e.key === 'o' || e.key === 'O') {
        toggleGridOverview();
      } else if (e.key === 'n' || e.key === 'N') {
        toggleNotes();
      } else if (e.key === 'Escape') {
        document.getElementById('grid-modal').classList.add('hidden');
        document.getElementById('image-modal').classList.add('hidden');
        document.getElementById('lab-modal').classList.add('hidden');
      }
    });

    // Toggle Grid Overview
    function toggleGridOverview() {
      const modal = document.getElementById('grid-modal');
      const isHidden = modal.classList.contains('hidden');
      if (isHidden) {
        renderGridCards();
        modal.classList.remove('hidden');
      } else {
        modal.classList.add('hidden');
      }
    }

    function renderGridCards() {
      const container = document.getElementById('grid-cards-container');
      container.innerHTML = '';
      titles.forEach((t, idx) => {
        const card = document.createElement('div');
        const isCurrent = (idx === currentSlide);
        card.className = `p-3.5 rounded-xl border text-xs cursor-pointer transition transform hover:scale-105 flex flex-col justify-between h-28 ${
          isCurrent 
            ? 'bg-cyan-950/80 border-cyan-400 shadow-lg shadow-cyan-500/20' 
            : 'bg-slate-900/80 border-slate-700 hover:border-slate-500 hover:bg-slate-800'
        }`;
        card.innerHTML = `
          <div>
            <span class="font-mono text-cyan-400 font-bold block mb-1">PAGE ${String(idx + 1).padStart(2, '0')}</span>
            <span class="text-slate-200 font-semibold leading-tight line-clamp-2">${t}</span>
          </div>
          <span class="text-[10px] text-slate-500 mt-2 block">${slides[idx].dataset.category || ''}</span>
        `;
        card.onclick = () => jumpToSlide(idx);
        container.appendChild(card);
      });
    }

    // Toggle Speaker Notes
    function toggleNotes() {
      const notes = document.getElementById('speaker-notes');
      notes.classList.toggle('hidden');
    }

    // Toggle Fullscreen
    function toggleFullscreen() {
      if (!document.fullscreenElement) {
        document.documentElement.requestFullscreen().catch(err => {
          console.log(`Fullscreen error: ${err.message}`);
        });
      } else {
        document.exitFullscreen();
      }
    }

    // Image Zoom Lightbox
    function zoomImage(src) {
      const modal = document.getElementById('image-modal');
      const img = document.getElementById('modal-img');
      img.src = src;
      modal.classList.remove('hidden');
    }

    function closeImageZoom() {
      document.getElementById('image-modal').classList.add('hidden');
    }

    // Interactive Lab Simulation
    function toggleInteractiveLab() {
      const modal = document.getElementById('lab-modal');
      modal.classList.toggle('hidden');
    }

    function loadSample(type) {
      const input = document.getElementById('lab-input');
      if (type === 'blank') {
        input.value = "   EMCU20491823   ";
        applyFormula('trim_upper');
      } else if (type === 'vessel') {
        input.value = "EVER GIVEN / 0124-055W";
        applyFormula('split_vessel');
      } else if (type === 'weight') {
        input.value = " 18,450.50 KGS ";
        applyFormula('clean_weight');
      }
    }

    function applyFormula(type) {
      const input = document.getElementById('lab-input').value;
      const fName = document.getElementById('lab-formula-name');
      const output = document.getElementById('lab-output');

      if (type === 'trim_upper') {
        fName.textContent = "=UPPER(TRIM(A2))";
        output.textContent = input.trim().toUpperCase();
      } else if (type === 'split_vessel') {
        fName.textContent = "=TRIM(LEFT(A2, FIND(\"/\", A2&\"/\")-1))";
        const parts = input.split('/');
        output.textContent = "船名: " + (parts[0] ? parts[0].trim() : '') + " | 航次: " + (parts[1] ? parts[1].trim() : '');
      } else if (type === 'clean_weight') {
        fName.textContent = "=VALUE(SUBSTITUTE(SUBSTITUTE(I2, \"KGS\", \"\"), \",\", \"\"))";
        const cleaned = input.toUpperCase().replace('KGS', '').replace(/,/g, '').trim();
        const num = parseFloat(cleaned);
        output.textContent = isNaN(num) ? "錯誤: 非數值" : num.toLocaleString('en-US', {minimumFractionDigits: 2}) + " (浮點數)";
      }
    }

    // Initialize
    updateSlide();
  </script>
</body>
</html>
"""

# Replace image placeholders with Base64 data URLs
for fname in fig_names:
    placeholder = f"IMAGE_{fname}"
    if fname in images_base64:
        html_template = html_template.replace(placeholder, images_base64[fname])
    else:
        html_template = html_template.replace(placeholder, "")

# Save output to both user workspace and artifact directory
out_workspace_path = "c:\\proj\\w12\\航運行政_Excel資料前處理_簡報投影片.html"
with open(out_workspace_path, "w", encoding="utf-8") as f:
    f.write(html_template)
print(f"Saved interactive presentation to: {out_workspace_path}")

out_artifact_dir = "C:\\Users\\user\\.gemini\\antigravity\\brain\\211eb1b6-505a-45ed-9361-b51c5d1e47cc"
out_artifact_path = os.path.join(out_artifact_dir, "航運行政_Excel資料前處理_簡報投影片.html")
with open(out_artifact_path, "w", encoding="utf-8") as f:
    f.write(html_template)
print(f"Saved interactive presentation to artifact: {out_artifact_path}")
