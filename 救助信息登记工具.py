import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import Workbook
from openpyxl.styles import Font
from datetime import datetime
import json
import os

SAVE_FILE = "save_data.json"

# 身份证解析
def parse_id_card(id_str):
    id_str = id_str.strip()
    if len(id_str) == 18:
        try:
            birth_year = int(id_str[6:10])
            birth_month = int(id_str[10:12])
            birth_day = int(id_str[12:14])
            sex_code = int(id_str[16])
            gender = "女" if sex_code % 2 == 0 else "男"
            birth_date = datetime(birth_year, birth_month, birth_day)
            today = datetime.now()
            age = today.year - birth_date.year
            if (today.month, today.day) < (birth_date.month, birth_date.day):
                age -= 1
            return gender, age
        except:
            return None, None
    elif len(id_str) == 15:
        try:
            birth_year = int("19" + id_str[6:8])
            birth_month = int(id_str[8:10])
            birth_day = int(id_str[10:12])
            sex_code = int(id_str[14])
            gender = "女" if sex_code % 2 == 0 else "男"
            birth_date = datetime(birth_year, birth_month, birth_day)
            today = datetime.now()
            age = today.year - birth_date.year
            if (today.month, today.day) < (birth_date.month, birth_date.day):
                age -= 1
            return gender, age
        except:
            return None, None
    return None, None

# 爱心基金阶梯计算
def calc_love_money(final_pay):
    try:
        val = float(final_pay)
    except:
        return 0
    if val >= 50000:
        return 10000
    elif val >= 40000:
        return 9000
    elif val >= 35000:
        return 8000
    elif val >= 30000:
        return 7000
    elif val >= 25000:
        return 6000
    elif val >= 20000:
        return 5000
    elif val >= 15000:
        return 4000
    elif val >= 10000:
        return 3000
    elif val >= 5000:
        return 2000
    else:
        return 0

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("救助信息登记工具（三项目）")
        self.root.geometry("1050x650")

        # 项目定义
        self.project_def = {
            "wucheng": {
                "name": "婺城助困助学助残",
                "headers": ["序号","姓名","性别","民族","年龄","身份证号码","家庭地址","联系电话","救助类型","是否脱贫户、低保户等","建议救助金额","备注"],
                "auto_money": 2000
            },
            "love": {
                "name": "爱心慈善基金（重大/一般疾病）",
                "headers": ["序号","姓名","性别","民族","年龄","身份证号码","家庭地址","联系电话","救助类型","是否脱贫户、低保户等","总花费","自费情况","民政局救助情况","最终审核金额自费情况","建议救助金额","备注"]
            },
            "shande": {
                "name": "善德关爱基金",
                "headers": ["序号","姓名","性别","民族","年龄","身份证号码","家庭地址","联系电话","救助类型","是否脱贫户、低保户等","损失情况","建议资助金额","备注"]
            }
        }
        self.project_id = tk.StringVar(value="wucheng")
        self.data = {
            "wucheng": [],
            "love": [],
            "shande": []
        }

        # 加载历史数据
        self.load_data()

        self.build_top()
        self.build_input()
        self.build_btn_bar()
        self.build_table()
        self.on_project_change()

    def save_data(self):
        """自动保存数据到json文件"""
        with open(SAVE_FILE, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    def load_data(self):
        """启动时读取保存的数据"""
        if os.path.exists(SAVE_FILE):
            try:
                with open(SAVE_FILE, "r", encoding="utf-8") as f:
                    self.data = json.load(f)
            except Exception:
                # 文件损坏就重置为空
                self.data = {"wucheng": [], "love": [], "shande": []}

    def build_top(self):
        f = ttk.LabelFrame(self.root, text="选择项目")
        f.pack(padx=10, pady=5, fill="x")
        ttk.Radiobutton(f, text="婺城助困助学助残", variable=self.project_id, value="wucheng", command=self.on_project_change).grid(row=0, column=0, padx=10)
        ttk.Radiobutton(f, text="爱心慈善基金（疾病）", variable=self.project_id, value="love", command=self.on_project_change).grid(row=0, column=1, padx=10)
        ttk.Radiobutton(f, text="善德关爱基金", variable=self.project_id, value="shande", command=self.on_project_change).grid(row=0, column=2, padx=10)

    def build_input(self):
        self.frame_input = ttk.LabelFrame(self.root, text="录入信息")
        self.frame_input.pack(padx=10, pady=5, fill="x")
        self.var_name = tk.StringVar()
        self.var_nation = tk.StringVar()
        self.var_id = tk.StringVar()
        self.var_addr = tk.StringVar()
        self.var_phone = tk.StringVar()
        self.var_help_type = tk.StringVar()
        self.var_is_low = tk.StringVar()
        self.var_total_cost = tk.StringVar()
        self.var_self_pay = tk.StringVar()
        self.var_mz_pay = tk.StringVar()
        self.var_lose = tk.StringVar()
        self.var_money = tk.StringVar()

        # 基础行1
        ttk.Label(self.frame_input, text="姓名").grid(row=0, column=0, padx=4, pady=4)
        ttk.Entry(self.frame_input, textvariable=self.var_name, width=12).grid(row=0, column=1, padx=4, pady=4)
        ttk.Label(self.frame_input, text="民族").grid(row=0, column=2, padx=4, pady=4)
        ttk.Entry(self.frame_input, textvariable=self.var_nation, width=10).grid(row=0, column=3, padx=4, pady=4)
        ttk.Label(self.frame_input, text="身份证号").grid(row=0, column=4, padx=4, pady=4)
        ttk.Entry(self.frame_input, textvariable=self.var_id, width=20).grid(row=0, column=5, padx=4, pady=4)

        # 基础行2
        ttk.Label(self.frame_input, text="家庭地址").grid(row=1, column=0, padx=4, pady=4)
        ttk.Entry(self.frame_input, textvariable=self.var_addr, width=24).grid(row=1, column=1, columnspan=2, padx=4, pady=4)
        ttk.Label(self.frame_input, text="联系电话").grid(row=1, column=3, padx=4, pady=4)
        ttk.Entry(self.frame_input, textvariable=self.var_phone, width=14).grid(row=1, column=4, padx=4, pady=4)
        ttk.Label(self.frame_input, text="救助类型").grid(row=1, column=5, padx=4, pady=4)
        ttk.Combobox(self.frame_input, textvariable=self.var_help_type, width=12, values=["助困","助学","助残","重大疾病","一般疾病","其他"]).grid(row=1, column=6, padx=4, pady=4)

        # 基础行3
        ttk.Label(self.frame_input, text="是否脱贫/低保").grid(row=2, column=0, padx=4, pady=4)
        ttk.Combobox(self.frame_input, textvariable=self.var_is_low, width=10, values=["是","否"]).grid(row=2, column=1, padx=4, pady=4)

        # ===== 爱心项目控件（总花费、自费、民政局救助）=====
        self.label_total = ttk.Label(self.frame_input, text="总花费")
        self.entry_total = ttk.Entry(self.frame_input, textvariable=self.var_total_cost, width=10)
        self.label_selfpay = ttk.Label(self.frame_input, text="自费情况")
        self.entry_selfpay = ttk.Entry(self.frame_input, textvariable=self.var_self_pay, width=10)
        self.label_mz = ttk.Label(self.frame_input, text="民政局救助")
        self.entry_mz = ttk.Entry(self.frame_input, textvariable=self.var_mz_pay, width=10)

        # ===== 善德项目控件（损失情况）=====
        self.label_lose = ttk.Label(self.frame_input, text="损失情况")
        self.entry_lose = ttk.Entry(self.frame_input, textvariable=self.var_lose, width=22)

        # 金额和备注
        ttk.Label(self.frame_input, text="建议救助金额").grid(row=3, column=0, padx=4, pady=4)
        vcmd = self.root.register(self.check_num)
        ttk.Entry(self.frame_input, textvariable=self.var_money, width=12, validate="key", validatecommand=(vcmd, "%P")).grid(row=3, column=1, padx=4, pady=4)
        ttk.Label(self.frame_input, text="备注").grid(row=3, column=2, padx=4, pady=4)
        self.txt_note = tk.Text(self.frame_input, width=28, height=2)
        self.txt_note.grid(row=3, column=3, columnspan=4, padx=4, pady=4)

    def check_num(self, inp):
        if inp == "":
            return True
        try:
            float(inp)
            return True
        except:
            return False

    def build_btn_bar(self):
        f = ttk.Frame(self.root)
        f.pack(pady=5)
        ttk.Button(f, text="添加本条", command=self.add_row).grid(row=0, column=0, padx=8)
        ttk.Button(f, text="删除选中", command=self.del_selected).grid(row=0, column=1, padx=8)
        ttk.Button(f, text="生成Excel", command=self.export_excel).grid(row=0, column=2, padx=8)
        ttk.Button(f, text="清空本项目", command=self.clear_current).grid(row=0, column=3, padx=8)

    def build_table(self):
        f = ttk.LabelFrame(self.root, text="已录入列表预览")
        f.pack(padx=10, pady=5, fill="both", expand=True)
        self.tree = ttk.Treeview(f, show="headings")
        self.tree.pack(fill="both", expand=True)
        self.scroll_y = ttk.Scrollbar(f, orient="vertical", command=self.tree.yview)
        self.scroll_y.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=self.scroll_y.set)

    def on_project_change(self):
        pid = self.project_id.get()
        # 全部先隐藏
        self.label_total.grid_remove()
        self.entry_total.grid_remove()
        self.label_selfpay.grid_remove()
        self.entry_selfpay.grid_remove()
        self.label_mz.grid_remove()
        self.entry_mz.grid_remove()
        self.label_lose.grid_remove()
        self.entry_lose.grid_remove()

        if pid == "love":
            # 爱心：显示总花费、自费、民政局救助，不显示损失
            self.label_total.grid(row=2, column=2, padx=4, pady=4)
            self.entry_total.grid(row=2, column=3, padx=4, pady=4)
            self.label_selfpay.grid(row=2, column=4, padx=4, pady=4)
            self.entry_selfpay.grid(row=2, column=5, padx=4, pady=4)
            self.label_mz.grid(row=2, column=6, padx=4, pady=4)
            self.entry_mz.grid(row=2, column=7, padx=4, pady=4)
        elif pid == "shande":
            # 善德：只显示损失情况
            self.label_lose.grid(row=2, column=2, padx=4, pady=4)
            self.entry_lose.grid(row=2, column=3, columnspan=3, padx=4, pady=4)
        # wucheng 婺城：这一行全部空着，啥都不显示

        headers = self.project_def[pid]["headers"]
        self.tree["columns"] = headers
        for item in self.tree.get_children():
            self.tree.delete(item)
        for c in headers:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=90, anchor="w")
        self.refresh_table()

    def refresh_table(self):
        pid = self.project_id.get()
        data_list = self.data[pid].copy()
        if pid == "love":
            def sort_key(row):
                try:
                    return -float(row[13])
                except:
                    return 0
            data_list.sort(key=sort_key)
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, row in enumerate(data_list, start=1):
            row[0] = idx
            self.tree.insert("", tk.END, values=row)

    def add_row(self):
        pid = self.project_id.get()
        name = self.var_name.get().strip()
        nation = self.var_nation.get().strip()
        idcard = self.var_id.get().strip()
        addr = self.var_addr.get().strip()
        phone = self.var_phone.get().strip()
        help_type = self.var_help_type.get().strip()
        is_low = self.var_is_low.get().strip()
        note = self.txt_note.get("1.0", tk.END).strip()

        if not name or not idcard:
            messagebox.showwarning("提醒", "姓名、身份证号不能为空！")
            return
        gender, age = parse_id_card(idcard)
        if gender is None or age is None:
            messagebox.showerror("错误", "身份证号码格式错误！")
            return

        if pid == "wucheng":
            money = self.var_money.get().strip()
            if not money:
                money = "2000"
            row = [0, name, gender, nation, age, idcard, addr, phone, help_type, is_low, money, note]
        elif pid == "love":
            total_cost = self.var_total_cost.get().strip()
            selfpay = self.var_self_pay.get().strip()
            mzpay = self.var_mz_pay.get().strip()
            try:
                s = float(selfpay) if selfpay else 0
                m = float(mzpay) if mzpay else 0
                final_pay = max(s - m, 0)
            except:
                final_pay = 0
            suggest_money = calc_love_money(final_pay)
            row = [0, name, gender, nation, age, idcard, addr, phone, help_type, is_low, total_cost, selfpay, mzpay, f"{final_pay:.2f}", f"{suggest_money}", note]
        elif pid == "shande":
            lose = self.var_lose.get().strip()
            money = self.var_money.get().strip()
            row = [0, name, gender, nation, age, idcard, addr, phone, help_type, is_low, lose, money, note]
        self.data[pid].append(row)
        self.refresh_table()
        self.save_data()  # 添加后自动保存
        # 清空输入框
        self.var_name.set("")
        self.var_nation.set("")
        self.var_id.set("")
        self.var_addr.set("")
        self.var_phone.set("")
        self.var_help_type.set("")
        self.var_is_low.set("")
        self.var_total_cost.set("")
        self.var_self_pay.set("")
        self.var_mz_pay.set("")
        self.var_lose.set("")
        self.var_money.set("")
        self.txt_note.delete("1.0", tk.END)

    def del_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("提醒", "请选中一行！")
            return
        pid = self.project_id.get()
        for item in sel:
            idx = self.tree.index(item)
            del self.data[pid][idx]
        self.refresh_table()
        self.save_data() # 删除后自动保存

    def clear_current(self):
        pid = self.project_id.get()
        if messagebox.askyesno("确认", "确定清空当前项目全部数据吗？此操作不可恢复！"):
            self.data[pid].clear()
            self.refresh_table()
            self.save_data() # 清空后自动保存

    def export_excel(self):
        pid = self.project_id.get()
        raw_list = self.data[pid].copy()
        if len(raw_list) == 0:
            messagebox.showwarning("提醒", "当前项目没有录入数据！")
            return
        if pid == "love":
            def sort_key(row):
                try:
                    return -float(row[13])
                except:
                    return 0
            export_list = sorted(raw_list, key=sort_key)
        else:
            export_list = raw_list
        for i, r in enumerate(export_list, start=1):
            r[0] = i
        wb = Workbook()
        ws = wb.active
        proj_info = self.project_def[pid]
        ws.title = proj_info["name"]
        headers = proj_info["headers"]
        ws.append(headers)
        for r in export_list:
            ws.append(r)
        bold_font = Font(bold=True)
        for col in range(1, len(headers)+1):
            cell = ws.cell(row=1, column=col)
            cell.font = bold_font
        fname = f"{proj_info['name']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        try:
            wb.save(fname)
            messagebox.showinfo("成功", f"文件已生成：{fname}")
        except Exception as e:
            messagebox.showerror("失败", str(e))

def main():
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
