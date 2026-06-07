#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
學力測驗 PDF 題庫抽取腳本：
讀取指定的 PDF 檔案，將其內容轉換為乾淨的純文字，並存檔於指定目錄。

用法：
    python scripts/extract_pdf.py --pdf "c:/2026Antigravity0606/近五年歷屆的學力測驗試題/114年國小五年級數學科學力測驗題本.pdf"
    python scripts/extract_pdf.py --folder "c:/2026Antigravity0606/近五年歷屆的學力測驗試題" --out-dir "c:/2026Antigravity0606/GEMS/output"
"""

import argparse
import os
import re
import sys
from pathlib import Path
from pypdf import PdfReader

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def clean_text(text: str) -> str:
    """清理 PDF 抽出的文字，去除多餘的頁碼、頁首頁尾等噪訊。"""
    # 去除常見的頁首/頁尾雜訊（例如：114年國小五年級學力測驗、第 X 頁...）
    text = re.sub(r"\d+年國小.*學力測驗.*題本", "", text)
    text = re.sub(r"第\s*\d+\s*頁\s*/\s*共\s*\d+\s*頁", "", text)
    text = re.sub(r"共\s*\d+\s*頁\s*第\s*\d+\s*頁", "", text)
    
    # 合併多個換行
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_single_pdf(pdf_path: Path, out_txt_path: Path) -> bool:
    """提取單個 PDF 的文字並存成 txt。"""
    if not pdf_path.exists():
        print(f"錯誤：找不到檔案 {pdf_path}", file=sys.stderr)
        return False

    try:
        print(f"正在讀取 PDF：{pdf_path.name} ...")
        reader = PdfReader(str(pdf_path))
        full_text = []

        for i, page in enumerate(reader.pages):
            text = page.extract_text()
            if text:
                full_text.append(f"--- [第 {i+1} 頁] ---")
                full_text.append(clean_text(text))

        result_text = "\n\n".join(full_text)
        
        # 建立父目錄
        out_txt_path.parent.mkdir(parents=True, exist_ok=True)
        out_txt_path.write_text(result_text, encoding="utf-8")
        print(f"成功匯出文字至：{out_txt_path}")
        return True
    except Exception as e:
        print(f"提取 {pdf_path.name} 失敗：{e}", file=sys.stderr)
        return False


def main() -> None:
    parser = argparse.ArgumentParser(description="學力測驗 PDF 題庫抽取腳本")
    parser.add_argument("--pdf", default=None, help="單個 PDF 檔案路徑")
    parser.add_argument("--folder", default=None, help="PDF 資料夾路徑")
    parser.add_argument("--out-dir", default=None, help="文字輸出目錄（預設輸出至 output/）")
    args = parser.parse_args()

    # 決定輸出目錄
    if args.out_dir:
        out_dir = Path(args.out_dir)
    else:
        # 預設輸出到專案的 output 目錄
        out_dir = Path(__file__).resolve().parent.parent / "output"
    
    out_dir.mkdir(parents=True, exist_ok=True)

    if args.pdf:
        pdf_path = Path(args.pdf)
        out_txt_path = out_dir / f"{pdf_path.stem}_extracted.txt"
        extract_single_pdf(pdf_path, out_txt_path)
    elif args.folder:
        folder_path = Path(args.folder)
        if not folder_path.is_dir():
            print(f"錯誤：找不到目錄 {folder_path}", file=sys.stderr)
            return
        
        pdf_files = list(folder_path.glob("*.pdf"))
        print(f"找到 {len(pdf_files)} 個 PDF 檔案，開始批次處理...")
        success_count = 0
        for f in pdf_files:
            out_txt_path = out_dir / f"{f.stem}_extracted.txt"
            if extract_single_pdf(f, out_txt_path):
                success_count += 1
        print(f"批次處理完成！成功：{success_count}/{len(pdf_files)}")
    else:
        # 預設處理近五年學測題本目錄
        default_folder = Path("c:/2026Antigravity0606/近五年歷屆的學力測驗試題")
        if default_folder.is_dir():
            print(f"未指定參數，預設處理目錄：{default_folder}")
            pdf_files = list(default_folder.glob("*.pdf"))
            for f in pdf_files:
                out_txt_path = out_dir / f"{f.stem}_extracted.txt"
                extract_single_pdf(f, out_txt_path)
        else:
            parser.print_help()


if __name__ == "__main__":
    main()
