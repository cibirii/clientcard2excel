#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
clientcard2excel — 微信聊天记录媒体解压工具
=========================================
从微信聊天记录导出的 zip（或已解压目录）中提取名片 / 合影 / 微信资料截图等媒体文件，
并打印清单，便于后续逐张读取、提取联系方式。

用法：
  python extract_media.py <zip或目录路径> [--out <输出目录>]

说明：
  - 微信导出 zip 内通常含「聊天记录.txt」与「聊天记录内的图片、视频和文件/」目录。
  - 已知 bug：微信有时只导出文字、照片未打入包（zip 内 0 图）。脚本会检测并告警，
    此时请让对方在微信里「逐张转发名片/合影」或「合并转发整段聊天」后重新导出。
"""

import os
import sys
import zipfile

IMG_EXTS = (".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif", ".heic")
MEDIA_EXTS = IMG_EXTS + (".mp4", ".mov", ".pdf")


def list_media(dir_path):
    files = []
    for root, _, names in os.walk(dir_path):
        for n in sorted(names):
            if n.lower().endswith(MEDIA_EXTS):
                files.append(os.path.join(root, n))
    return files


def extract_zip(zip_path, out_dir):
    with zipfile.ZipFile(zip_path, "r") as z:
        names = z.namelist()
        media = [n for n in names if n.lower().endswith(MEDIA_EXTS)]
        txts = [n for n in names if n.lower().endswith(".txt")]
        os.makedirs(out_dir, exist_ok=True)
        for n in media + txts:
            z.extract(n, out_dir)
        return len(media), len(txts), len(names)


def main():
    if len(sys.argv) < 2:
        print("用法: extract_media.py <zip或目录> [--out <dir>]")
        sys.exit(1)
    src = sys.argv[1]
    out = None
    if "--out" in sys.argv:
        out = sys.argv[sys.argv.index("--out") + 1]

    if os.path.isdir(src):
        files = list_media(src)
        print(f"目录: {src}")
        print(f"媒体文件数: {len(files)}")
        for f in files:
            print("  ", os.path.relpath(f, src))
        if not files:
            print("⚠️ 未发现媒体文件，请确认路径或微信导出是否包含图片。")
        return

    if zipfile.is_zipfile(src):
        base = out or (os.path.splitext(src)[0] + "_media")
        n_media, n_txt, n_total = extract_zip(src, base)
        print(f"ZIP: {src}")
        print(f"解压到: {base}")
        print(f"媒体文件: {n_media}  文本文件(txt): {n_txt}  ZIP内总条目: {n_total}")
        if n_media == 0:
            print("\n⚠️ ZIP 内 0 张图片！多为微信导出时照片未打入包（已知 bug）。")
            print("   请让对方重新在微信里「逐张转发名片/合影」或「合并转发整段聊天」后重新导出。")
        files = list_media(base)
        for f in files:
            print("  ", os.path.basename(f))
        return

    print("❌ 无法识别的路径（需 zip 文件或目录）")
    sys.exit(1)


if __name__ == "__main__":
    main()
