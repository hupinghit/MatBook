import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from openpyxl import load_workbook

# 读取 Excel 文件
def load_materials_from_excel(file_path):
    materials = {}
    try:
        workbook = load_workbook(filename=file_path)
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            materials[sheet_name] = {}
            header = [cell.value for cell in sheet[1]]  # 第一行为属性名称
            for row in sheet.iter_rows(min_row=2, values_only=True):  # 从第二行开始解析数据
                material_name = row[0]  # 第一列为材料名称
                material_properties = {}
                for i in range(1, len(row)):  # 从第二列开始解析属性
                    material_properties[header[i]] = row[i]
                materials[sheet_name][material_name] = material_properties
        return materials
    except Exception as e:
        messagebox.showerror("错误", f"无法读取 Excel 文件: {e}")
        return None

# 点击材料名称时的回调函数
def show_material_properties(event):
    selected_material = material_listbox.get(material_listbox.curselection())
    material = find_material_by_name(selected_material, materials)
    if material:
        result = "\n".join([f"{key}: {value}" for key, value in material.items()])
    else:
        result = f"材料 '{selected_material}' 没有定义！"
    result_label.config(text=result)

# 查找材料属性
def find_material_by_name(material_name, materials):
    for category, items in materials.items():
        if material_name in items:
            return items[material_name]
    return None

# 加载材料数据
FILE_PATH = "materials.xlsx"  # Excel 文件路径
materials = load_materials_from_excel(FILE_PATH)
if not materials:
    exit()

# 创建主窗口
root = tk.Tk()
root.title("材料属性查询")

# 创建左侧材料列表框架
left_frame = ttk.Frame(root, padding="10")
left_frame.grid(row=0, column=0, sticky=(tk.W, tk.N, tk.S))

# 材料列表标题
material_label = ttk.Label(left_frame, text="材料列表")
material_label.grid(row=0, column=0, sticky=tk.W)

# 材料列表框
material_listbox = tk.Listbox(left_frame, width=30, height=15)
material_listbox.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# 将所有材料名称添加到列表框中
for category, items in materials.items():
    for material_name in items.keys():
        material_listbox.insert(tk.END, material_name)

# 绑定点击事件
material_listbox.bind("<<ListboxSelect>>", show_material_properties)

# 创建右侧属性显示框架
right_frame = ttk.Frame(root, padding="10")
right_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))

# 属性显示区域
result_label = ttk.Label(right_frame, text="请点击左侧材料名称查看属性", wraplength=300)
result_label.grid(row=0, column=0, sticky=(tk.W, tk.E))

# 启动主循环
root.mainloop()
