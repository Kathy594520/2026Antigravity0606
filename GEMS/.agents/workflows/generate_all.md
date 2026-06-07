---
name: generate_all
description: 一鍵執行完整升級流程，讀取 input/ 的材料，依據指令自動派發並執行對應的 Skill，最後將生成成果匯出至 output/ 目錄中。
---

# /generate_all

請為使用者執行完整的一鍵自動化工作流：

1. **偵測與導向**：
   - 檢查 `input/` 資料夾中的輸入檔案類型與配置。
   - 若 `input/` 為空，請提示使用者放入材料。

2. **Skills 分發與執行**：
   - **文字或題目材料** ➔ 調用 `chinese_exam_generator` 進行 108 課綱學力測驗題目生成，並呼叫 `extract_pdf.py` 作為 PDF 題庫比對參考，產出練習題 Markdown 及 Word。
   - **教學課文或寫作主題** ➔ 調用 `worksheet_visualizer` 進行 PIRLS 六大層次或寫作「起承轉合」學習單設計，並自動透過 **Nano Banana Pro** 生生圖。
   - **影音 URL 或音訊檔** ➔ 調用 `video_subtitle_helper` 下載、提取並生成 SRT 字幕及 docx 摘要。

3. **收工回報**：
   - 條列報告本次處理了哪些 input 材料。
   - 顯示所有新產出檔案的完整路徑（位於 `output/` 目錄），方便使用者直接點擊開啟。
