#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GEMS Bazi Assets Generator
同時生成：
1. PDF 報告書 (A4 格式，內置五行長條圖，支援中文)
2. HTML 互動式簡報 (毛玻璃感、漸層流光、含五行模擬與流年滑桿)
3. Obsidian 筆記 (.md 格式，包含 YAML frontmatter)
4. bazi_data.json (供前端 Web Dashboard 讀取與動態渲染)
支援單人模式與雙人合盤模式。
"""

import os
import sys
import json
import argparse
from fpdf import FPDF

# 確保控制台輸出編碼正確
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# 顏色定義
COLORS = {
    "水": (2, 132, 199),   # 藍
    "木": (5, 150, 105),   # 綠
    "火": (220, 38, 38),   # 紅
    "土": (124, 58, 237),  # 紫
    "金": (245, 158, 11),  # 黃
}

class BaziPDF(FPDF):
    def __init__(self, title_text="個人命盤天賦報告"):
        super().__init__()
        self.title_text = title_text

    def header(self):
        # 頁首
        self.set_font("msjh", size=9)
        self.set_text_color(120, 120, 120)
        self.cell(0, 10, self.title_text, align="R", new_x="LMARGIN", new_y="NEXT")
        self.ln(5)

    def footer(self):
        # 頁尾
        self.set_y(-15)
        self.set_font("msjh", size=9)
        self.set_text_color(120, 120, 120)
        self.cell(0, 10, f"第 {self.page_no()} 頁", align="C", new_x="LMARGIN", new_y="NEXT")

def draw_pdf_profile(pdf, data, is_partner=False):
    # 確保起點在最左側
    pdf.set_x(pdf.l_margin)
    
    # 大標題或小標題
    title_color = COLORS.get(data["daymaster"][0], (0, 102, 153))
    pdf.set_font("msjh", size=16)
    pdf.set_text_color(*title_color)
    prefix = "合作夥伴：" if is_partner else "主命主："
    pdf.cell(pdf.epw, 10, f"{prefix}{data['name']} 的命盤分析", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)

    # 基本資訊區
    pdf.set_font("msjh", size=10.5)
    pdf.set_text_color(50, 50, 50)
    pdf.set_fill_color(248, 250, 252)
    
    gender_str = "男" if data.get("gender") == "男" else "女"
    info_text = (
        f"【基本資訊】\n"
        f"• 姓名：{data['name']} ({gender_str})\n"
        f"• 出生時間：{data['birth_time']} (農曆：{data.get('birth_lunar', '未提供')})\n"
        f"• 命主日元：{data['daymaster']}\n"
        f"• 格局特徵：{data.get('profile_features', '未提供')}"
    )
    pdf.multi_cell(pdf.epw, 7.5, text=info_text, border=1, fill=True)
    pdf.ln(5)
    pdf.set_x(pdf.l_margin)

    # 五行能量區
    pdf.set_font("msjh", size=12.5)
    pdf.set_text_color(*title_color)
    pdf.cell(pdf.epw, 8, "五行能量數據分析", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(1)
    
    pdf.set_font("msjh", size=9.5)
    pdf.set_text_color(80, 80, 80)
    
    max_bar_width = 100
    # 尋找最高分數以作比例尺
    scores = [el.get("score", 0) for el in data.get("elements", [])]
    max_score = max(scores) if scores else 100
    if max_score == 0:
        max_score = 100
        
    for el in data.get("elements", []):
        pdf.set_x(pdf.l_margin)
        el_name = el["name"]
        el_role = el.get("role", "")
        el_score = el.get("score", 0)
        
        pdf.cell(50, 8, text=f"{el_name} ({el_role}): {el_score}分")
        
        # 繪製能量條背景
        x = pdf.get_x()
        y = pdf.get_y() + 1.5
        pdf.set_fill_color(230, 230, 230)
        pdf.rect(x, y, max_bar_width, 4.5, "F")
        
        # 繪製能量條填充
        if el_score > 0:
            fill_width = (el_score / float(max_score)) * max_bar_width
            pdf.set_fill_color(*COLORS.get(el_name, (100, 100, 100)))
            pdf.rect(x, y, fill_width, 4.5, "F")
        else:
            pdf.set_font("msjh", size=8)
            pdf.set_text_color(150, 150, 150)
            pdf.text(x + 5, y + 3.5, "0分 (無先天能量)")
            pdf.set_font("msjh", size=9.5)
            pdf.set_text_color(80, 80, 80)
        
        pdf.ln(7.5)
        
    pdf.ln(5)
    pdf.set_x(pdf.l_margin)
    
    # 解析區
    pdf.set_font("msjh", size=12.5)
    pdf.set_text_color(*title_color)
    pdf.cell(pdf.epw, 8, "[解析]", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(1)
    
    pdf.set_font("msjh", size=10)
    pdf.set_text_color(50, 50, 50)
    
    for i, para in enumerate(data.get("analysis", [])):
        pdf.set_x(pdf.l_margin)
        para_title = para.get("title", f"解析 {i+1}")
        para_text = para.get("text", "")
        pdf.set_font("msjh", "B", size=10)
        pdf.cell(pdf.epw, 6.5, text=para_title, new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("msjh", size=10)
        pdf.multi_cell(pdf.epw, 6.5, text=para_text)
        pdf.ln(3)
        
    pdf.ln(2)
    pdf.set_x(pdf.l_margin)

    # 人生指引區
    pdf.set_font("msjh", size=12.5)
    pdf.set_text_color(*title_color)
    pdf.cell(pdf.epw, 8, "[人生指引]", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(1)
    
    pdf.set_font("msjh", size=10)
    pdf.set_text_color(50, 50, 50)
    
    for guide in data.get("guidelines", []):
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(pdf.epw, 6.5, text=f"• {guide}")
    pdf.ln(5)
    pdf.set_x(pdf.l_margin)
    
    # 覺察提問
    if data.get("questions"):
        pdf.set_fill_color(245, 245, 245)
        pdf.set_font("msjh", "B", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(pdf.epw, 7, text="【自我覺察提問】", new_x="LMARGIN", new_y="NEXT", fill=True)
        pdf.set_font("msjh", size=10)
        for i, q in enumerate(data.get("questions", [])):
            pdf.set_x(pdf.l_margin)
            pdf.multi_cell(pdf.epw, 6.5, text=f"{i+1}. {q}", fill=True)
        pdf.ln(5)
        pdf.set_x(pdf.l_margin)

def build_html_presentation(data, output_html_path):
    """
    生成毛玻璃、流光色彩、極致美感的 HTML 互動式簡報
    """
    has_partner = "partner" in data
    
    # 建構 HTML 內容
    html_title = f"{data['name']} - 個人命盤天賦簡報" if not has_partner else f"{data['name']} & {data['partner']['name']} - 合作契合度命盤簡報"
    
    # 將資料序列化為 JSON 供前端 JS 使用
    bazi_json_str = json.dumps(data, ensure_ascii=False)
    
    html_content = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{html_title}</title>
    <!-- Google Fonts -->
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;800&family=Noto+Sans+TC:wght@300;400;700;900&display=swap" rel="stylesheet">
    <style>
        :root {{
            --primary: #1245;
            --font-main: 'Outfit', 'Noto Sans TC', sans-serif;
            --bg-dark: #0f172a;
            --card-bg: rgba(255, 255, 255, 0.06);
            --card-border: rgba(255, 255, 255, 0.12);
            --text-light: #f8fafc;
            --text-muted: #94a3b8;
            --water: #0284c7;
            --wood: #059669;
            --fire: #dc2626;
            --earth: #7c3aed;
            --metal: #f59e0b;
        }}

        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }}

        body {{
            font-family: var(--font-main);
            background-color: var(--bg-dark);
            color: var(--text-light);
            min-height: 100vh;
            overflow: hidden;
            display: flex;
            justify-content: center;
            align-items: center;
            position: relative;
        }}

        /* 背景流光流體效果 */
        .bg-glow {{
            position: absolute;
            width: 600px;
            height: 600px;
            border-radius: 50%;
            background: radial-gradient(circle, rgba(99, 102, 241, 0.15) 0%, rgba(0,0,0,0) 70%);
            filter: blur(80px);
            z-index: -1;
            pointer-events: none;
            animation: float 20s infinite alternate;
        }}
        .bg-glow-1 {{
            top: -10%;
            left: -10%;
            background: radial-gradient(circle, rgba(14, 165, 233, 0.15) 0%, rgba(0,0,0,0) 70%);
        }}
        .bg-glow-2 {{
            bottom: -10%;
            right: -10%;
            background: radial-gradient(circle, rgba(168, 85, 247, 0.15) 0%, rgba(0,0,0,0) 70%);
        }}

        @keyframes float {{
            0% {{ transform: translate(0, 0) scale(1); }}
            100% {{ transform: translate(50px, 50px) scale(1.1); }}
        }}

        /* 簡報主容器 */
        .presentation-container {{
            width: 90vw;
            height: 85vh;
            max-width: 1200px;
            max-height: 750px;
            background: rgba(15, 23, 42, 0.6);
            backdrop-filter: blur(20px);
            -webkit-backdrop-filter: blur(20px);
            border: 1px solid var(--card-border);
            border-radius: 24px;
            box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
            display: flex;
            flex-direction: column;
            position: relative;
            z-index: 10;
        }}

        /* 頁首導覽 */
        .header {{
            padding: 24px 40px;
            border-bottom: 1px solid var(--card-border);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        .brand {{
            font-size: 20px;
            font-weight: 800;
            background: linear-gradient(135deg, #38bdf8, #818cf8);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }}

        .slide-counter {{
            font-size: 14px;
            color: var(--text-muted);
            letter-spacing: 2px;
        }}

        /* 幻燈片主視窗 */
        .slides-viewport {{
            flex: 1;
            position: relative;
            overflow: hidden;
        }}

        .slide {{
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            padding: 40px 60px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            opacity: 0;
            transform: translateX(100px);
            transition: all 0.6s cubic-bezier(0.4, 0, 0.2, 1);
            pointer-events: none;
        }}

        .slide.active {{
            opacity: 1;
            transform: translateX(0);
            pointer-events: auto;
        }}

        .slide.prev-slide {{
            opacity: 0;
            transform: translateX(-100px);
        }}

        /* 排版樣式 */
        .slide-title {{
            font-size: 32px;
            font-weight: 900;
            margin-bottom: 24px;
            display: flex;
            align-items: center;
            gap: 12px;
        }}

        .slide-grid {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 40px;
            align-items: start;
        }}

        /* 玻璃面板 */
        .glass-panel {{
            background: var(--card-bg);
            border: 1px solid var(--card-border);
            border-radius: 16px;
            padding: 24px;
        }}

        .glass-panel h3 {{
            margin-bottom: 16px;
            font-size: 18px;
            color: #38bdf8;
        }}

        p {{
            line-height: 1.8;
            font-size: 16px;
            color: #cbd5e1;
            margin-bottom: 12px;
        }}

        ul {{
            list-style: none;
        }}

        li {{
            margin-bottom: 12px;
            line-height: 1.6;
            font-size: 15px;
            color: #e2e8f0;
            position: relative;
            padding-left: 20px;
        }}

        li::before {{
            content: "✦";
            position: absolute;
            left: 0;
            color: #38bdf8;
        }}

        /* 能量條視覺 */
        .bar-container {{
            margin-bottom: 16px;
        }}

        .bar-label {{
            display: flex;
            justify-content: space-between;
            margin-bottom: 6px;
            font-size: 14px;
        }}

        .bar-bg {{
            width: 100%;
            height: 10px;
            background: rgba(255,255,255,0.1);
            border-radius: 5px;
            overflow: hidden;
        }}

        .bar-fill {{
            height: 100%;
            border-radius: 5px;
            width: 0;
            transition: width 1s cubic-bezier(0.4, 0, 0.2, 1);
        }}

        /* 控制滑桿區 */
        .slider-control {{
            margin-top: 20px;
            padding: 16px;
            background: rgba(0,0,0,0.2);
            border-radius: 12px;
            border: 1px solid var(--card-border);
        }}
        
        .slider-row {{
            display: flex;
            align-items: center;
            gap: 15px;
            margin-bottom: 10px;
        }}

        .slider-row label {{
            width: 80px;
            font-size: 14px;
        }}

        .slider-row input[type="range"] {{
            flex: 1;
            accent-color: #38bdf8;
        }}

        /* 導覽按鍵 */
        .nav-controls {{
            padding: 24px 40px;
            border-top: 1px solid var(--card-border);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        .btn {{
            background: rgba(255, 255, 255, 0.08);
            border: 1px solid var(--card-border);
            color: var(--text-light);
            padding: 10px 24px;
            border-radius: 9999px;
            cursor: pointer;
            font-weight: 600;
            transition: all 0.3s ease;
            display: flex;
            align-items: center;
            gap: 8px;
        }}

        .btn:hover {{
            background: var(--text-light);
            color: var(--bg-dark);
            transform: translateY(-2px);
        }}

        .btn:disabled {{
            opacity: 0.3;
            cursor: not-allowed;
            transform: none;
        }}

        .btn-primary {{
            background: linear-gradient(135deg, #38bdf8, #818cf8);
            border: none;
        }}

        .btn-primary:hover {{
            box-shadow: 0 0 15px rgba(56, 189, 248, 0.5);
            color: white;
        }}

        /* 自訂樣式供合盤卡片 */
        .badge {{
            display: inline-block;
            padding: 4px 12px;
            border-radius: 9999px;
            font-size: 12px;
            font-weight: bold;
            background: rgba(255,255,255,0.15);
        }}

        .badge-water {{ background-color: var(--water); }}
        .badge-wood {{ background-color: var(--wood); }}
        .badge-fire {{ background-color: var(--fire); }}
        .badge-earth {{ background-color: var(--earth); }}
        .badge-metal {{ background-color: var(--metal); }}

    </style>
</head>
<body>
    <div class="bg-glow bg-glow-1"></div>
    <div class="bg-glow bg-glow-2"></div>

    <div class="presentation-container">
        <!-- 頁首 -->
        <div class="header">
            <div class="brand">GEMS Bazi Studio</div>
            <div class="slide-counter" id="slide-counter">SLIDE 1 / 5</div>
        </div>

        <!-- 投影片內容區 -->
        <div class="slides-viewport" id="slides-viewport">
            <!-- 投影片 1：歡迎與命主基本資訊 -->
            <div class="slide active" id="slide-0">
                <h2 class="slide-title">🔮 {data['name']} - 個人天賦命盤解析</h2>
                <div class="slide-grid">
                    <div class="glass-panel">
                        <h3>先天的原廠設定</h3>
                        <p><strong>性別：</strong> {data.get('gender', '女')}</p>
                        <p><strong>西元出生年月日時：</strong> {data['birth_time']}</p>
                        <p><strong>農曆時間：</strong> {data.get('birth_lunar', '未提供')}</p>
                        <p><strong>命主日元：</strong> <span class="badge badge-metal">{data['daymaster']}</span></p>
                        <p><strong>核心格局特徵：</strong> {data.get('profile_features', '載入中')}</p>
                    </div>
                    <div class="glass-panel">
                        <h3>天賦角色與心智設定</h3>
                        <p style="font-size: 15px; color: var(--text-muted)">
                            我們將您的八字命盤能量重新框架為現代心理學的角色模型，幫助您釐清內在潛力。
                        </p>
                        <ul style="margin-top: 15px;">
                            <li>核心心智：{data['daymaster'][0]}金日主之特質</li>
                            <li>天賦優勢：能量平衡下的特長</li>
                            <li>人生課題：需落地實踐的能量引導</li>
                        </ul>
                    </div>
                </div>
            </div>

            <!-- 投影片 2：五行與能量模擬器 -->
            <div class="slide" id="slide-1">
                <h2 class="slide-title">📊 五行能量分布與流年模擬器</h2>
                <div class="slide-grid">
                    <div class="glass-panel">
                        <h3>五行先天能量百分比</h3>
                        <div id="bar-chart-container">
                            <!-- JS 動態生成能量條 -->
                        </div>
                    </div>
                    <div class="glass-panel">
                        <h3>互動模擬器：流年能量調適</h3>
                        <p style="font-size: 14px; color: var(--text-muted); margin-bottom: 10px;">
                            拖動下方大運流年能量滑桿，模擬外部能量對先天八字帶來的影響：
                        </p>
                        <div class="slider-control">
                            <div class="slider-row">
                                <label>流年水 (食傷)</label>
                                <input type="range" id="sim-water" min="-50" max="100" value="0" oninput="simulateEnergy()">
                                <span id="val-water">+0</span>
                            </div>
                            <div class="slider-row">
                                <label>流年木 (財星)</label>
                                <input type="range" id="sim-wood" min="-50" max="100" value="0" oninput="simulateEnergy()">
                                <span id="val-wood">+0</span>
                            </div>
                            <div class="slider-row">
                                <label>流年火 (官殺)</label>
                                <input type="range" id="sim-fire" min="-50" max="100" value="0" oninput="simulateEnergy()">
                                <span id="val-fire">+0</span>
                            </div>
                        </div>
                        <p style="font-size: 13px; color: #38bdf8; margin-top: 15px;" id="simulation-advice">
                            提示：2026 丙午年引動火與土的能量，適合補水(運動與排解)與補木(實務產出)以獲得平衡。
                        </p>
                    </div>
                </div>
            </div>

            <!-- 投影片 3：深度特質解析 -->
            <div class="slide" id="slide-2">
                <h2 class="slide-title">🧠 內在性格與心理拉鋸解析</h2>
                <div class="slide-grid">
                    <div class="glass-panel" style="grid-column: span 2;">
                        <h3>[解析] 深度解讀</h3>
                        <div id="analysis-content" style="max-height: 380px; overflow-y: auto; padding-right: 10px;">
                            <!-- JS 動態渲染段落 -->
                        </div>
                    </div>
                </div>
            </div>

            <!-- 投影片 4：人生引導與具體實踐 -->
            <div class="slide" id="slide-3">
                <h2 class="slide-title">🌱 [人生指引] 能量落地的實踐方案</h2>
                <div class="slide-grid">
                    <div class="glass-panel">
                        <h3>具體行動指引</h3>
                        <ul id="guidelines-list" style="margin-top: 10px;">
                            <!-- JS 動態渲染 -->
                        </ul>
                    </div>
                    <div class="glass-panel">
                        <h3>💭 自我覺察問題</h3>
                        <div id="questions-list" style="margin-top: 10px;">
                            <!-- JS 動態渲染 -->
                        </div>
                    </div>
                </div>
            </div>

            <!-- 雙人合盤投影片 (如果有合盤數據) -->
            {"<div class=\"slide\" id=\"slide-4\">" if has_partner else ""}
                {"<h2 class=\"slide-title\">👥 雙人合作契合度與合盤分析</h2>" if has_partner else ""}
                {"<div class=\"slide-grid\">" if has_partner else ""}
                    {"<div class=\"glass-panel\">" if has_partner else ""}
                        {"<h3>雙方能量對比</h3>" if has_partner else ""}
                        {"<p><strong>主命主：</strong>" + data['name'] + " (" + data['daymaster'] + ")</p>" if has_partner else ""}
                        {"<p><strong>夥伴：</strong>" + (data['partner']['name'] if has_partner else '') + " (" + (data['partner']['daymaster'] if has_partner else '') + ")</p>" if has_partner else ""}
                        {"<div id=\"partner-bar-comparison\" style=\"margin-top: 15px;\"></div>" if has_partner else ""}
                    {"</div>" if has_partner else ""}
                    {"<div class=\"glass-panel\">" if has_partner else ""}
                        {"<h3>溝通避坑與合作指南</h3>" if has_partner else ""}
                        {"<p id=\"relation-analysis-text\" style=\"font-size:14px; line-height:1.7;\"></p>" if has_partner else ""}
                    {"</div>" if has_partner else ""}
                {"</div>" if has_partner else ""}
            {"</div>" if has_partner else ""}

            <!-- 最後一頁：免責聲明與下載 -->
            <div class="slide" id="slide-final">
                <h2 class="slide-title">✨ 踏上自我覺察之旅</h2>
                <div style="text-align: center; max-width: 700px; margin: 0 auto;">
                    <div class="glass-panel" style="margin-bottom: 40px; padding: 40px;">
                        <h3 style="font-size: 24px; color: #38bdf8; margin-bottom: 20px;">報告編譯成功</h3>
                        <p style="margin-bottom: 30px;">
                            本命盤測算資料已同步儲存於本機 output/ 目錄。您可以點擊對話中的連結下載完整 PDF 與 Obsidian 筆記檔案。
                        </p>
                        <div style="font-size: 13px; color: var(--text-muted); line-height: 1.6; border-top: 1px solid var(--card-border); padding-top: 20px;">
                            免責聲明：本報告所提供之命理分析與心理測算，旨在作為促進自我覺察與心靈成長之輔助工具。所有分析內容與建議僅供參考。
                        </div>
                    </div>
                </div>
            </div>

        </div>

        <!-- 底部導覽按鈕 -->
        <div class="nav-controls">
            <button class="btn" id="prev-btn" onclick="prevSlide()" disabled>
                ← Prev
            </button>
            <button class="btn btn-primary" id="next-btn" onclick="nextSlide()">
                Next →
            </button>
        </div>
    </div>

    <script>
        // 載入八字原始資料
        const baziData = {bazi_json_str};

        let currentSlideIndex = 0;
        const slides = document.querySelectorAll('.slide');
        const counter = document.getElementById('slide-counter');
        const prevBtn = document.getElementById('prev-btn');
        const nextBtn = document.getElementById('next-btn');

        function updateSlide() {{
            slides.forEach((slide, idx) => {{
                slide.classList.remove('active', 'prev-slide');
                if (idx === currentSlideIndex) {{
                    slide.classList.add('active');
                }} else if (idx < currentSlideIndex) {{
                    slide.classList.add('prev-slide');
                }}
            }});

            counter.textContent = `SLIDE ${{currentSlideIndex + 1}} / slide數量為 ${{slides.length}}`;
            prevBtn.disabled = currentSlideIndex === 0;
            if (currentSlideIndex === slides.length - 1) {{
                nextBtn.textContent = 'Finish';
            }} else {{
                nextBtn.innerHTML = 'Next &rarr;';
            }}
        }}

        function nextSlide() {{
            if (currentSlideIndex < slides.length - 1) {{
                currentSlideIndex++;
                updateSlide();
            }} else {{
                alert('感謝使用本簡報！檔案已安全留存於本機。');
            }}
        }}

        function prevSlide() {{
            if (currentSlideIndex > 0) {{
                currentSlideIndex--;
                updateSlide();
            }}
        }}

        // 監聽鍵盤方向鍵
        document.addEventListener('keydown', (e) => {{
            if (e.key === 'ArrowRight' || e.key === 'Space') {{
                nextSlide();
            }} else if (e.key === 'ArrowLeft') {{
                prevSlide();
            }}
        }});

        // 初始化圖表
        function initCharts() {{
            const container = document.getElementById('bar-chart-container');
            container.innerHTML = '';
            
            baziData.elements.forEach(el => {{
                const pct = (el.score / 240.0) * 100;
                let colorVar = 'var(--text-muted)';
                if (el.name === '水') colorVar = 'var(--water)';
                if (el.name === '木') colorVar = 'var(--wood)';
                if (el.name === '火') colorVar = 'var(--fire)';
                if (el.name === '土') colorVar = 'var(--earth)';
                if (el.name === '金') colorVar = 'var(--metal)';

                container.innerHTML += `
                    <div class="bar-container">
                        <div class="bar-label">
                            <span>${{el.name}} (${{el.role}})</span>
                            <span>${{el.score}} 分</span>
                        </div>
                        <div class="bar-bg">
                            <div class="bar-fill" style="width: ${{Math.min(pct, 100)}}%; background-color: ${{colorVar}}"></div>
                        </div>
                    </div>
                `;
            }});

            // 渲染解析內容
            const analysisDiv = document.getElementById('analysis-content');
            analysisDiv.innerHTML = '';
            baziData.analysis.forEach((para, idx) => {{
                analysisDiv.innerHTML += `
                    <div style="margin-bottom: 20px;">
                        <h4 style="color:#38bdf8; font-size:16px; margin-bottom:8px;">${{para.title}}</h4>
                        <p style="font-size:14px; line-height:1.7;">${{para.text}}</p>
                    </div>
                `;
            }});

            // 渲染指引與問題
            const guideUl = document.getElementById('guidelines-list');
            guideUl.innerHTML = '';
            baziData.guidelines.forEach(g => {{
                guideUl.innerHTML += `<li>${{g}}</li>`;
            }});

            const qDiv = document.getElementById('questions-list');
            qDiv.innerHTML = '';
            baziData.questions.forEach((q, idx) => {{
                qDiv.innerHTML += `
                    <div class="glass-panel" style="margin-bottom: 12px; padding: 12px; border-left: 4px solid #818cf8; background:rgba(0,0,0,0.15)">
                        <p style="font-size: 13.5px; margin: 0; color:#e2e8f0;">${{idx+1}}. ${{q}}</p>
                    </div>
                `;
            }});

            // 渲染雙人合盤對比 (如果有的話)
            const partnerComparison = document.getElementById('partner-bar-comparison');
            if (partnerComparison && baziData.partner) {{
                partnerComparison.innerHTML = '';
                // 比較表渲染
                baziData.elements.forEach((el, idx) => {{
                    const partnerEl = baziData.partner.elements.find(e => e.name === el.name) || {{score: 0}};
                    partnerComparison.innerHTML += `
                        <div style="margin-bottom: 10px; font-size:13px;">
                            <div style="display:flex; justify-content:space-between; margin-bottom:3px;">
                                <span>${{el.name}}</span>
                                <span>${{baziData.name}}: ${{el.score}}分 | ${{baziData.partner.name}}: ${{partnerEl.score}}分</span>
                            </div>
                            <div style="display:flex; gap:4px;">
                                <div style="flex:1; height:6px; background:rgba(255,255,255,0.05); border-radius:3px; overflow:hidden;">
                                    <div style="height:100%; width:${{Math.min(el.score/2.4, 100)}}%; background-color:#38bdf8;"></div>
                                </div>
                                <div style="flex:1; height:6px; background:rgba(255,255,255,0.05); border-radius:3px; overflow:hidden;">
                                    <div style="height:100%; width:${{Math.min(partnerEl.score/2.4, 100)}}%; background-color:#f43f5e;"></div>
                                </div>
                            </div>
                        </div>
                    `;
                }});

                // 關係分析
                document.getElementById('relation-analysis-text').textContent = baziData.relation_analysis || '分析合作默契、溝通模式與互補之處。';
            }}
        }}

        // 流年模擬器邏輯
        function simulateEnergy() {{
            const waterVal = parseInt(document.getElementById('sim-water').value);
            const woodVal = parseInt(document.getElementById('sim-wood').value);
            const fireVal = parseInt(document.getElementById('sim-fire').value);

            document.getElementById('val-water').textContent = (waterVal >= 0 ? '+' : '') + waterVal;
            document.getElementById('val-wood').textContent = (woodVal >= 0 ? '+' : '') + woodVal;
            document.getElementById('val-fire').textContent = (fireVal >= 0 ? '+' : '') + fireVal;

            // 重新計算能量百分比並更新 UI
            const container = document.getElementById('bar-chart-container');
            container.innerHTML = '';
            
            baziData.elements.forEach(el => {{
                let modifier = 0;
                if (el.name === '水') modifier = waterVal;
                if (el.name === '木') modifier = woodVal;
                if (el.name === '火') modifier = fireVal;

                const simulatedScore = Math.max(0, el.score + modifier);
                const pct = (simulatedScore / 240.0) * 100;
                
                let colorVar = 'var(--text-muted)';
                if (el.name === '水') colorVar = 'var(--water)';
                if (el.name === '木') colorVar = 'var(--wood)';
                if (el.name === '火') colorVar = 'var(--fire)';
                if (el.name === '土') colorVar = 'var(--earth)';
                if (el.name === '金') colorVar = 'var(--metal)';

                container.innerHTML += `
                    <div class="bar-container">
                        <div class="bar-label">
                            <span>${{el.name}} (${{el.role}})</span>
                            <span>${{simulatedScore}} 分 (先天 ${{el.score}} ${{modifier >= 0 ? '+' : ''}}${{modifier}})</span>
                        </div>
                        <div class="bar-bg">
                            <div class="bar-fill" style="width: ${{Math.min(pct, 100)}}%; background-color: ${{colorVar}}"></div>
                        </div>
                    </div>
                `;
            }});

            // 根據模擬結果給予簡短提示
            const advice = document.getElementById('simulation-advice');
            if (waterVal > 30) {{
                advice.textContent = '💡 提示：增加水能量有助於發揮創意才華（傷官配印），可緩解火（官殺）的壓力內耗。';
            }} else if (woodVal > 30) {{
                advice.textContent = '💡 提示：增加木能量有助於實務執行落地，改善天馬行空的規劃，讓創意實現。';
            }} else if (fireVal > 30) {{
                advice.textContent = '⚠️ 警示：火（官殺）能量過旺可能帶來過大的工作壓力與內在完美主義，需適當放鬆（補水）。';
            }} else {{
                advice.textContent = '提示：拖動滑桿模擬流年運勢對您的五行能量產生的微調影響，尋找最佳平衡方案。';
            }}
        }}

        // 載入完成後初始化
        window.onload = () => {{
            initCharts();
            updateSlide();
        }};
    </script>
</body>
</html>
"""
    with open(output_html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

def build_obsidian_markdown(data, output_md_path):
    """
    生成 Obsidian YAML 元數據格式筆記
    """
    has_partner = "partner" in data
    
    # 建立五行 YAML 對應
    elements_yaml = {}
    for el in data.get("elements", []):
        elements_yaml[el["name"]] = el["score"]
        
    yaml_header = {
        "type": "BaziProfile",
        "name": data["name"],
        "gender": data.get("gender", "女"),
        "birth_time": data["birth_time"],
        "daymaster": data["daymaster"],
        "features": data.get("profile_features", ""),
        "scores": elements_yaml,
        "date_analyzed": "2026-06-07"
    }
    
    if has_partner:
        partner_elements_yaml = {}
        for el in data["partner"].get("elements", []):
            partner_elements_yaml[el["name"]] = el["score"]
        yaml_header["partner"] = {
            "name": data["partner"]["name"],
            "gender": data["partner"].get("gender", "男"),
            "daymaster": data["partner"]["daymaster"],
            "scores": partner_elements_yaml
        }

    md_content = []
    md_content.append("---")
    # 簡單寫入 YAML 屬性
    for k, v in yaml_header.items():
        if isinstance(v, dict):
            md_content.append(f"{k}:")
            for sub_k, sub_v in v.items():
                md_content.append(f"  {sub_k}: {sub_v}")
        else:
            md_content.append(f"{k}: {v}")
    md_content.append("---")
    md_content.append("")
    md_content.append(f"# {data['name']} 的個人命盤天賦筆記")
    md_content.append("")
    md_content.append(f"**格局特徵**: {data.get('profile_features', '')}")
    md_content.append("")
    
    md_content.append("## 五行數據")
    for el in data.get("elements", []):
        md_content.append(f"- **{el['name']}** ({el.get('role', '')}): {el['score']} 分")
    md_content.append("")
    
    md_content.append("## [解析]")
    for para in data.get("analysis", []):
        md_content.append(f"### {para['title']}")
        md_content.append(para["text"])
        md_content.append("")
        
    md_content.append("## [人生指引]")
    for g in data.get("guidelines", []):
        md_content.append(f"- {g}")
    md_content.append("")
    
    md_content.append("## 自我覺察提問")
    for i, q in enumerate(data.get("questions", [])):
        md_content.append(f"{i+1}. {q}")
    md_content.append("")
    
    if has_partner:
        md_content.append("---")
        md_content.append(f"# 與 {data['partner']['name']} 的雙人合盤分析")
        md_content.append("")
        md_content.append(f"**夥伴命日元**: {data['partner']['daymaster']}")
        md_content.append("")
        md_content.append("## 關係分析與合作指南")
        md_content.append(data.get("relation_analysis", ""))
        md_content.append("")
        
    with open(output_md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md_content))

def main():
    parser = argparse.ArgumentParser(description="GEMS Bazi Assets Generator")
    parser.add_argument("json_file", help="Path to input JSON file containing Bazi details")
    args = parser.parse_args()
    
    if not os.path.exists(args.json_file):
        print(f"Error: Input file {args.json_file} does not exist.")
        sys.exit(1)
        
    with open(args.json_file, "r", encoding="utf-8") as f:
        data = json.load(f)
        
    name = data["name"]
    output_dir = "c:/2026Antigravity0606/GEMS/output"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    pdf_path = f"{output_dir}/{name}_個人命盤天賦報告.pdf"
    html_path = f"{output_dir}/{name}_個人命盤天賦簡報.html"
    md_path = f"{output_dir}/{name}_個人命盤天賦筆記.md"
    json_path = f"{output_dir}/bazi_data.json"
    
    # 1. 生成 PDF
    print(f"Generating PDF: {pdf_path}")
    pdf = BaziPDF(title_text=f"{name} - 個人命盤天賦與心理測算報告")
    pdf.alias_nb_pages()
    
    # 註冊中文字型
    font_path = "C:/Windows/Fonts/msjh.ttc"
    if not os.path.exists(font_path):
        # 尋找替代的中文字型
        font_path = "C:/Windows/Fonts/mingliu.ttc"
    if os.path.exists(font_path):
        pdf.add_font("msjh", "", font_path)
        pdf.add_font("msjh", "B", font_path)
    else:
        # 使用預設字型
        print("Warning: Chinese font not found, falling back to default Arial")
        pdf.add_font("msjh", "", "Arial")
        pdf.add_font("msjh", "B", "Arial")
        
    pdf.add_page()
    
    # 大標題
    title_color = COLORS.get(data["daymaster"][0], (0, 102, 153))
    pdf.set_font("msjh", size=20)
    pdf.set_text_color(*title_color)
    pdf.cell(pdf.epw, 15, f"{name} - 個人命盤天賦與心理測算報告", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(5)
    
    # 繪製主命盤
    draw_pdf_profile(pdf, data)
    
    # 如果有合盤夥伴，在第二頁繪製夥伴與關係分析
    if "partner" in data:
        pdf.add_page()
        draw_pdf_profile(pdf, data["partner"], is_partner=True)
        
        pdf.ln(5)
        pdf.set_x(pdf.l_margin)
        pdf.set_font("msjh", size=13)
        pdf.set_text_color(153, 51, 51)
        pdf.cell(pdf.epw, 10, "雙人合盤關係分析與合作指南", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(2)
        
        pdf.set_font("msjh", size=10)
        pdf.set_text_color(50, 50, 50)
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(pdf.epw, 7, text=data.get("relation_analysis", "無關係分析數據"))
        pdf.ln(5)
        
    # 免責聲明
    pdf.set_x(pdf.l_margin)
    pdf.set_font("msjh", size=8)
    pdf.set_text_color(150, 150, 150)
    disclaimer = (
        "免責聲明：本報告所提供之命理分析與心理測算，旨在作為促進自我覺察與心靈成長之輔助工具。所有分析內容與建議僅供參考，實際人生決策仍請依循您的個人自由意志與專業領域評估進行。"
    )
    pdf.multi_cell(pdf.epw, 5, text=disclaimer, align="C")
    
    pdf.output(pdf_path)
    print("✓ PDF generated successfully")
    
    # 2. 生成 HTML 互動簡報
    print(f"Generating HTML: {html_path}")
    build_html_presentation(data, html_path)
    print("✓ HTML slideshow generated successfully")
    
    # 3. 生成 Obsidian Markdown 筆記
    print(f"Generating MD: {md_path}")
    build_obsidian_markdown(data, md_path)
    print("✓ Markdown note generated successfully")
    
    # 4. 輸出 bazi_data.json
    print(f"Writing JSON data: {json_path}")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print("✓ bazi_data.json updated successfully")
    
    print("\n[SUCCESS] All Bazi assets generated successfully!")
    print(f"PDF: {pdf_path}")
    print(f"HTML: {html_path}")
    print(f"Markdown: {md_path}")
    print(f"JSON: {json_path}")

if __name__ == "__main__":
    main()
