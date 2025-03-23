import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox


def export_file_names_to_excel(directory, excel_path):
    """
    将指定目录下的文件名导出到 Excel 文件的 A 列
    """
    if not os.path.exists(directory):
        messagebox.showerror("错误", "目录不存在，请检查路径！")
        return

    # 确保 'input' 文件夹存在，如果不存在则创建
    input_folder = os.path.dirname(excel_path)  # 获取文件路径所在的文件夹
    if not os.path.exists(input_folder):
        os.makedirs(input_folder)

    # 获取目录下的文件和文件夹名称
    file_names = os.listdir(directory)

    # 创建 DataFrame
    df = pd.DataFrame({"oldName": file_names, "newName": ""})  # A 列为旧文件名，B 列为空
    df.to_excel(excel_path, index=False, sheet_name="重命名")

    messagebox.showinfo("成功", f"文件名已导出到 {excel_path} 的 A 列，请填写新文件名到 B 列。")



def rename_files_from_excel(excel_path, base_folder):
    """
    根据 Excel 文件内容批量重命名文件或文件夹
    """
    # 读取 Excel 文件
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        messagebox.showerror("错误", f"无法读取 Excel 文件：{e}")
        return

    # 检查是否包含所需的列
    if "oldName" not in df.columns or "newName" not in df.columns:
        messagebox.showerror("错误", "Excel 文件必须包含 'oldName' 和 'newName' 列！")
        return

    # 确保列的数据类型为字符串
    df["oldName"] = df["oldName"].astype(str)
    df["newName"] = df["newName"].astype(str)

    # 遍历 Excel 表格的每一行，先将文件重命名为临时名称
    temp_names = {}
    for _, row in df.iterrows():
        old_name = row["oldName"]
        new_name = row["newName"]

        if pd.isna(new_name) or new_name.strip() == "":
            continue

        old_path = os.path.join(base_folder, old_name)
        temp_path = os.path.join(base_folder, f"temp_{old_name}")

        # 使用临时名称避免冲突
        if os.path.exists(old_path):
            os.rename(old_path, temp_path)
            temp_names[temp_path] = new_name

    # 再将临时名称改为目标名称
    for temp_path, new_name in temp_names.items():
        new_path = os.path.join(base_folder, new_name)

        try:
            os.rename(temp_path, new_path)
        except Exception as e:
            messagebox.showwarning("警告", f"重命名失败：{temp_path} -> {new_name}\n错误：{e}")

    messagebox.showinfo("完成", "重命名操作完成！")



def select_directory():
    """
    选择目录并导出文件名
    """
    directory = filedialog.askdirectory(title="选择目录")
    if not directory:
        return

    excel_file = "./input/rename.xlsx"

    # 导出文件名到 Excel
    export_file_names_to_excel(directory, excel_file)

    # 存储用户选择的目录
    global selected_directory
    selected_directory = directory


def perform_rename():
    """
    执行重命名
    """
    excel_file = "./input/rename.xlsx"

    # 检查是否选择了目录
    if not selected_directory:
        messagebox.showerror("错误", "请先选择一个目录并导出文件名！")
        return

    # 检查是否存在 Excel 文件
    if not os.path.exists(excel_file):
        messagebox.showerror("错误", f"找不到文件：{excel_file}")
        return

    # 执行重命名
    rename_files_from_excel(excel_file, selected_directory)


def main():
    """
    主 GUI 界面
    """
    root = tk.Tk()
    root.title("批量文件重命名工具")

    # 窗口大小
    root.geometry("400x200")
    root.resizable(False, False)

    # 导出文件名按钮
    btn_export = tk.Button(
        root, text="导出文件名到 Excel", command=select_directory, width=30, height=2
    )
    btn_export.pack(pady=20)

    # 执行重命名按钮
    btn_rename = tk.Button(
        root, text="执行文件重命名", command=perform_rename, width=30, height=2
    )
    btn_rename.pack(pady=20)

    # 运行主循环
    root.mainloop()


# 存储用户选择的目录
selected_directory = ""

if __name__ == "__main__":
    main()
