import os
import hashlib
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook

def get_file_md5(file_path, chunk_size=4096):
    hash_obj = hashlib.md5()
    try:
        with open(file_path, 'rb') as f:
            while chunk := f.read(chunk_size):
                hash_obj.update(chunk)
        return hash_obj.hexdigest()
    except Exception:
        return None

def find_duplicates(folder):
    size_dict = defaultdict(list)
    for root, _, files in os.walk(folder):
        for fname in files:
            fpath = os.path.join(root, fname)
            try:
                size = os.path.getsize(fpath)
                if size == 0:
                    continue
                size_dict[size].append(fpath)
            except Exception:
                continue

    dup_dict = defaultdict(list)
    for size, paths in size_dict.items():
        if len(paths) < 2:
            continue
        for p in paths:
            md5 = get_file_md5(p)
            if md5:
                dup_dict[md5].append(p)

    return {h: ps for h, ps in dup_dict.items() if len(ps) >= 2}

def export_to_excel(dup_map, save_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "重复文件"
    ws.append(["重复组号", "文件路径", "文件大小(MB)", "MD5"])

    group_no = 1
    for h, paths in dup_map.items():
        for p in paths:
            try:
                size_mb = round(os.path.getsize(p) / 1024 / 1024, 2)
            except:
                size_mb = "读取失败"
            ws.append([group_no, p, size_mb, h])
        group_no += 1

    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 70
    ws.column_dimensions['C'].width = 16
    ws.column_dimensions['D'].width = 36

    wb.save(save_path)

def main():
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title="选择要扫描重复文件的文件夹")

    if not folder:
        print("未选择文件夹，退出")
        return

    print(f"正在扫描: {folder}")
    duplicates = find_duplicates(folder)

    if not duplicates:
        print("✅ 未发现重复文件")
        return

    excel_path = os.path.join(os.path.expanduser("~"), "重复文件清单.xlsx")
    export_to_excel(duplicates, excel_path)

    print(f"\n📊 扫描完成！Excel 已保存到：")
    print(excel_path)
    print(f"共 {len(duplicates)} 组重复文件")

if __name__ == "__main__":
    main()