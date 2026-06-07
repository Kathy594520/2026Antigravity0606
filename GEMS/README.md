# Gemini Gems 升級為 Antigravity Agent 工作流專案

這個專案是從您原先在 Google Gemini 上的 Gems 升級整併而來的本地端 Agent 工作流專案。專案已針對 **Google Antigravity** 平台進行優化。

## 📁 專案結構

```text
GEMS/
├── README.md             # 專案說明文件
├── .agents/              # Agent 設定與技能目錄
│   ├── skills/           # 核心技能（Skills）定義
│   │   ├── chinese_exam_generator/   # 國語科命題與審題系統
│   │   ├── worksheet_visualizer/     # 圖文學習單與生圖生成器
│   │   └── video_subtitle_helper/    # 影片字幕與摘要助手
│   └── workflows/        # 自動化工作流（Workflows）
│       └── generate_all.md
├── input/                # 每次執行的浮動輸入材料置放處
├── output/               # 執行後的輸出成品（Word、SRT、圖片等）
└── scripts/              # 本地端執行的輔助 Python 腳本
    └── extract_pdf.py    # 學力測驗 PDF 題庫抽取腳本
```

## 🛠️ 三大核心技能說明

1. **國語科命題與審題系統 (`chinese_exam_generator`)**
   - **用途**：結合 108 課綱與歷屆考題，為小學 3~6 年級國語科出題、仿題及審題。
   - **本機特色**：可直接讀取並解析 `c:\2026Antigravity0606\近五年歷屆的學力測驗試題` 目錄下的歷屆學測題本 PDF。
   - **產出**：高品質試題與詳解，並支援匯出為 Word (`.docx`) 檔。

2. **教學引導與圖文學習單生成器 (`worksheet_visualizer`)**
   - **用途**：自動將課文轉化為 Bloom 六大認知層次的圖文學習單設計，或作文引導多格插畫版面。
   - **本機特色**：支援將生圖指示傳遞給 Antigravity 內建的 **Nano Banana Pro** 生圖引擎。

3. **影片字幕與摘要助手 (`video_subtitle_helper`)**
   - **用途**：下載影片、提取音訊、自動生成字幕及摘要。
   - **本機特色**：串接本地安裝之 `yt-dlp` 與 `ffmpeg`。

## 🚀 使用方式

1. 將您的材料放置於 `input/` 資料夾（如：課文文字、音檔、PDF 檔案等）。
2. 在 Antigravity Chat 中，輸入 `/[workflow-名稱]` 或是指示載入 `.agents/skills` 下對應的技能即可開始。
3. 產生的結果（例如 Word 文件、字幕檔、生圖等）將自動存放於 `output/` 資料夾。
