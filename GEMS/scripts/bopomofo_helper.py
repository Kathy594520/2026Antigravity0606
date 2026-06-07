#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GEMS Bopomofo Helper
自動將中文字元標註注音，輸出為 HTML <ruby> 標籤格式，適合國小低中年級教材預覽。
"""

import sys
import os

try:
    import pypinyin
except ImportError:
    print("Warning: pypinyin not installed. Running 'pip install pypinyin' first.")

# 確保控制台編碼為 UTF-8
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

def is_chinese_char(c):
    """
    判斷字元是否為中文字 (CJK)
    """
    if len(c) != 1:
        return False
    return '\u4e00' <= c <= '\u9fff'

def convert_to_ruby(text):
    """
    將文字轉換為 HTML ruby 標籤格式
    """
    output = []
    
    for char in text:
        if is_chinese_char(char):
            try:
                # 逐字獲取注音，避免因標點符號對齊錯位
                res = pypinyin.pinyin(char, style=pypinyin.Style.BOPOMOFO, errors='default')
                if res and res[0] and res[0][0]:
                    bopo = res[0][0].replace(" ", "")
                    output.append(f"<ruby>{char}<rt>{bopo}</rt></ruby>")
                else:
                    output.append(char)
            except Exception as e:
                print(f"Error calling pypinyin for {char}: {e}", file=sys.stderr)
                output.append(char)
        else:
            # 非中文字元直接輸出
            if char == '\n':
                output.append("<br>\n")
            elif char == '\r':
                pass
            else:
                output.append(char)
                
    return "".join(output)

def make_html_preview(ruby_text, title="國語試題/教材注音預覽"):
    """
    生成完整的 HTML 預覽頁面
    """
    html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <title>{title}</title>
    <style>
        body {{
            font-family: "Noto Sans TC", "Microsoft JhengHei", sans-serif;
            background-color: #f8fafc;
            color: #1e293b;
            padding: 40px;
            display: flex;
            justify-content: center;
        }}
        .preview-card {{
            background: white;
            padding: 40px;
            border-radius: 16px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
            max-width: 800px;
            width: 100%;
            line-height: 3.5; /* 給注音足夠的高度 */
            font-size: 24px;
        }}
        h2 {{
            line-height: 1.5;
            margin-bottom: 24px;
            border-bottom: 2px solid #3b82f6;
            padding-bottom: 12px;
            font-size: 28px;
        }}
        ruby {{
            ruby-position: over;
            margin: 0 2px;
        }}
        rt {{
            font-size: 11px;
            color: #64748b;
            user-select: none;
        }}
    </style>
</head>
<body>
    <div class="preview-card">
        <h2>{title}</h2>
        <div class="content">
            {ruby_text}
        </div>
    </div>
</body>
</html>
"""
    return html

def main():
    if len(sys.argv) < 2:
        print("Usage: python bopomofo_helper.py <text_to_convert_or_file_path>")
        print("Or:    python bopomofo_helper.py --test")
        sys.exit(1)

    input_arg = sys.argv[1]
    
    if input_arg == "--test":
        # 測試模式
        test_text = "床前明月光，疑是地上霜。舉頭望明月，低頭思故鄉。"
        ruby_res = convert_to_ruby(test_text)
        print("Test input: ", test_text)
        print("Test output: ", ruby_res)
        
        output_dir = "c:/2026Antigravity0606/GEMS/output"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
        test_html_path = f"{output_dir}/bopomofo_test.html"
        with open(test_html_path, "w", encoding="utf-8") as f:
            f.write(make_html_preview(ruby_res, "靜夜思注音預覽"))
        print(f"✓ Test HTML preview generated: {test_html_path}")
        return

    # 判斷是否為檔案路徑
    if os.path.exists(input_arg):
        with open(input_arg, "r", encoding="utf-8") as f:
            text = f.read()
    else:
        text = input_arg

    ruby_text = convert_to_ruby(text)
    
    # 預設寫入 output 供 Dashboard 讀取
    output_dir = "c:/2026Antigravity0606/GEMS/output"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    html_path = f"{output_dir}/worksheet_bopomofo.html"
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(make_html_preview(ruby_text))
        
    print(f"✓ Success! Ruby text generated and saved to: {html_path}")
    # 同時印出 ruby 格式字串
    print(ruby_text)

if __name__ == "__main__":
    main()
