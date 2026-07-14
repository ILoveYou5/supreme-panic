# supreme-panic
# 试着使用gayhub / get familiar with github
# 会上传一些自己日常学习代码及编写代码的心得 / I'll upload my insights on learning and writing code
# 下面是最近用claude code生成的一段python代码，功能是调用一个程序，对一个文件夹下的zip文件进行处理，生成的文件放到另外一个文件夹 / Below is a python program with created by claude code , the program can use a .exe file to handle zip files in a directory , output files in another directory.
## 以下开始 program #1 / here begins program #1
#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
批量解密工具
自动处理目录下所有 ZIP 文件
"""

import os
import subprocess
import sys
import json
from pathlib import Path
from tkinter import Tk, Text, Scrollbar, Button, Frame, Label, Entry
from tkinter import messagebox, filedialog
import threading

class DecryptGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("ZIP 文件批量解密工具")
        self.root.geometry("850x650")

        # 工作目录和配置文件
        self.work_dir = Path(__file__).parent
        self.config_file = self.work_dir / "config.json"
        self.decrypt_tool_path = None
        self.input_dir = None
        self.output_dir = None

        # 加载配置
        self.load_config()

        # 标题
        title_label = Label(root, text="ZIP 文件批量解密工具", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)

        # === 配置区域 ===
        config_container = Frame(root)
        config_container.pack(pady=5, padx=10, fill="x")

        # 解密工具路径配置
        tool_frame = Frame(config_container)
        tool_frame.pack(fill="x", pady=3)

        Label(tool_frame, text="解密工具路径:", font=("Arial", 10), width=15, anchor="w").pack(side="left", padx=5)
        self.tool_path_entry = Entry(tool_frame, font=("Arial", 9), state="readonly")
        self.tool_path_entry.pack(side="left", fill="x", expand=True, padx=5)
        Button(tool_frame, text="选择文件", command=self.select_decrypt_tool,
               bg="#2196F3", fg="white", font=("Arial", 9, "bold"), width=10).pack(side="left", padx=5)

        # 输入目录配置
        input_frame = Frame(config_container)
        input_frame.pack(fill="x", pady=3)

        Label(input_frame, text="待解密文件目录:", font=("Arial", 10), width=15, anchor="w").pack(side="left", padx=5)
        self.input_dir_entry = Entry(input_frame, font=("Arial", 9), state="readonly")
        self.input_dir_entry.pack(side="left", fill="x", expand=True, padx=5)
        Button(input_frame, text="选择目录", command=self.select_input_dir,
               bg="#2196F3", fg="white", font=("Arial", 9, "bold"), width=10).pack(side="left", padx=5)

        # 输出目录配置
        output_frame = Frame(config_container)
        output_frame.pack(fill="x", pady=3)

        Label(output_frame, text="解密后文件目录:", font=("Arial", 10), width=15, anchor="w").pack(side="left", padx=5)
        self.output_dir_entry = Entry(output_frame, font=("Arial", 9), state="readonly")
        self.output_dir_entry.pack(side="left", fill="x", expand=True, padx=5)
        Button(output_frame, text="选择目录", command=self.select_output_dir,
               bg="#2196F3", fg="white", font=("Arial", 9, "bold"), width=10).pack(side="left", padx=5)

        # 更新所有路径显示
        self.update_all_path_displays()

        # 信息标签
        info_label = Label(root, text="配置好路径后，点击下方按钮开始批量解密", font=("Arial", 10))
        info_label.pack(pady=8)

        # 按钮框架
        button_frame = Frame(root)
        button_frame.pack(pady=10)

        # 开始按钮
        self.start_button = Button(button_frame, text="开始解密", command=self.start_decrypt,
                                   bg="#4CAF50", fg="white", font=("Arial", 12, "bold"),
                                   width=15, height=2)
        self.start_button.pack(side="left", padx=5)

        # 退出按钮
        exit_button = Button(button_frame, text="退出", command=self.root.quit,
                           bg="#f44336", fg="white", font=("Arial", 12, "bold"),
                           width=15, height=2)
        exit_button.pack(side="left", padx=5)

        # 日志显示区域
        log_frame = Frame(root)
        log_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # 滚动条
        scrollbar = Scrollbar(log_frame)
        scrollbar.pack(side="right", fill="y")

        # 文本框
        self.log_text = Text(log_frame, wrap="word", yscrollcommand=scrollbar.set,
                            font=("Consolas", 9))
        self.log_text.pack(fill="both", expand=True)
        scrollbar.config(command=self.log_text.yview)

        # 显示初始配置信息
        self.log("=" * 80)
        self.log("当前配置:")
        self.log(f"  解密工具: {self.decrypt_tool_path or '未配置'}")
        self.log(f"  输入目录: {self.input_dir or '未配置'}")
        self.log(f"  输出目录: {self.output_dir or '未配置'}")
        self.log("=" * 80)

    def log(self, message):
        """添加日志到文本框"""
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.root.update()

    def load_config(self):
        """加载配置文件"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)

                    # 加载解密工具路径
                    tool_path = config.get('decrypt_tool_path', '')
                    if tool_path and Path(tool_path).exists():
                        self.decrypt_tool_path = tool_path
                    else:
                        # 尝试默认路径
                        default_path = self.work_dir / "LiptsDecryptTool.exe"
                        if default_path.exists():
                            self.decrypt_tool_path = str(default_path)

                    # 加载输入目录
                    input_dir = config.get('input_dir', '')
                    if input_dir and Path(input_dir).exists():
                        self.input_dir = input_dir
                    else:
                        # 默认为工作目录
                        self.input_dir = str(self.work_dir)

                    # 加载输出目录
                    output_dir = config.get('output_dir', '')
                    if output_dir and Path(output_dir).exists():
                        self.output_dir = output_dir
                    else:
                        # 默认为工作目录下的"解密文件"
                        self.output_dir = str(self.work_dir / "解密文件")
            else:
                # 首次运行，设置默认值
                default_tool = self.work_dir / "LiptsDecryptTool.exe"
                if default_tool.exists():
                    self.decrypt_tool_path = str(default_tool)

                self.input_dir = str(self.work_dir)
                self.output_dir = str(self.work_dir / "解密文件")

                # 保存默认配置
                self.save_config()
        except Exception as e:
            print(f"加载配置失败: {e}")
            # 设置默认值
            self.input_dir = str(self.work_dir)
            self.output_dir = str(self.work_dir / "解密文件")

    def save_config(self):
        """保存配置文件"""
        try:
            config = {
                'decrypt_tool_path': self.decrypt_tool_path,
                'input_dir': self.input_dir,
                'output_dir': self.output_dir
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置失败: {e}")

    def update_all_path_displays(self):
        """更新所有路径显示"""
        # 更新解密工具路径
        self.tool_path_entry.config(state="normal")
        self.tool_path_entry.delete(0, "end")
        if self.decrypt_tool_path:
            self.tool_path_entry.insert(0, self.decrypt_tool_path)
        else:
            self.tool_path_entry.insert(0, "未配置 - 请点击右侧按钮选择")
        self.tool_path_entry.config(state="readonly")

        # 更新输入目录
        self.input_dir_entry.config(state="normal")
        self.input_dir_entry.delete(0, "end")
        if self.input_dir:
            self.input_dir_entry.insert(0, self.input_dir)
        else:
            self.input_dir_entry.insert(0, "未配置 - 请点击右侧按钮选择")
        self.input_dir_entry.config(state="readonly")

        # 更新输出目录
        self.output_dir_entry.config(state="normal")
        self.output_dir_entry.delete(0, "end")
        if self.output_dir:
            self.output_dir_entry.insert(0, self.output_dir)
        else:
            self.output_dir_entry.insert(0, "未配置 - 请点击右侧按钮选择")
        self.output_dir_entry.config(state="readonly")

    def select_decrypt_tool(self):
        """选择解密工具路径"""
        initial_dir = self.work_dir
        if self.decrypt_tool_path:
            initial_dir = Path(self.decrypt_tool_path).parent

        file_path = filedialog.askopenfilename(
            title="选择 LiptsDecryptTool.exe",
            initialdir=initial_dir,
            filetypes=[("可执行文件", "*.exe"), ("所有文件", "*.*")]
        )

        if file_path:
            self.decrypt_tool_path = file_path
            self.save_config()
            self.update_all_path_displays()
            self.log(f"\n已设置解密工具路径: {self.decrypt_tool_path}")

    def select_input_dir(self):
        """选择输入目录"""
        initial_dir = self.input_dir if self.input_dir else self.work_dir

        dir_path = filedialog.askdirectory(
            title="选择待解密文件所在目录",
            initialdir=initial_dir
        )

        if dir_path:
            self.input_dir = dir_path
            self.save_config()
            self.update_all_path_displays()
            self.log(f"\n已设置输入目录: {self.input_dir}")

    def select_output_dir(self):
        """选择输出目录"""
        initial_dir = self.output_dir if self.output_dir else self.work_dir

        dir_path = filedialog.askdirectory(
            title="选择解密后文件保存目录",
            initialdir=initial_dir
        )

        if dir_path:
            self.output_dir = dir_path
            self.save_config()
            self.update_all_path_displays()
            self.log(f"\n已设置输出目录: {self.output_dir}")

    def start_decrypt(self):
        """开始解密过程"""
        self.start_button.config(state="disabled")
        thread = threading.Thread(target=self.decrypt_all_files)
        thread.daemon = True
        thread.start()

    def decrypt_all_files(self):
        """批量解密所有文件"""
        try:
            # 检查是否配置了解密工具路径
            if not self.decrypt_tool_path:
                self.log("\n错误: 未配置解密工具路径！")
                messagebox.showerror("错误", "未配置解密工具路径！\n\n请点击 '选择文件' 按钮选择 LiptsDecryptTool.exe")
                self.start_button.config(state="normal")
                return

            decrypt_tool = Path(self.decrypt_tool_path)
            if not decrypt_tool.exists():
                self.log(f"\n错误: 解密工具不存在: {decrypt_tool}")
                messagebox.showerror("错误", f"解密工具不存在:\n{decrypt_tool}\n\n请重新设置路径")
                self.start_button.config(state="normal")
                return

            # 检查输入目录
            if not self.input_dir:
                self.log("\n错误: 未配置输入目录！")
                messagebox.showerror("错误", "未配置输入目录！\n\n请点击 '选择目录' 按钮设置待解密文件所在目录")
                self.start_button.config(state="normal")
                return

            input_path = Path(self.input_dir)
            if not input_path.exists():
                self.log(f"\n错误: 输入目录不存在: {input_path}")
                messagebox.showerror("错误", f"输入目录不存在:\n{input_path}\n\n请重新设置路径")
                self.start_button.config(state="normal")
                return

            # 检查输出目录
            if not self.output_dir:
                self.log("\n错误: 未配置输出目录！")
                messagebox.showerror("错误", "未配置输出目录！\n\n请点击 '选择目录' 按钮设置解密后文件保存目录")
                self.start_button.config(state="normal")
                return

            output_path = Path(self.output_dir)
            # 确保输出目录存在
            output_path.mkdir(parents=True, exist_ok=True)
            self.log(f"\n输出目录已准备: {output_path}")

            # 查找所有 ZIP 文件
            zip_files = list(input_path.glob("*.zip"))

            if not zip_files:
                self.log(f"\n未找到 ZIP 文件！(搜索目录: {input_path})")
                messagebox.showwarning("警告", f"在以下目录未找到 ZIP 文件:\n{input_path}")
                self.start_button.config(state="normal")
                return

            self.log(f"\n找到 {len(zip_files)} 个 ZIP 文件")
            self.log(f"输入目录: {input_path}")
            self.log(f"输出目录: {output_path}")
            self.log("=" * 80)

            success_count = 0
            fail_count = 0

            # 处理每个文件
            for idx, zip_file in enumerate(zip_files, 1):
                self.log(f"\n[{idx}/{len(zip_files)}] 正在处理: {zip_file.name}")

                output_file = output_path / zip_file.name

                try:
                    # 调用解密工具
                    result = subprocess.run(
                        [str(decrypt_tool), str(zip_file), str(output_file)],
                        capture_output=True,
                        text=True,
                        encoding='utf-8',
                        errors='ignore'
                    )

                    # 显示输出
                    if result.stdout:
                        for line in result.stdout.strip().split('\n'):
                            if line.strip():
                                self.log(f"  {line}")

                    # 检查是否成功
                    if result.returncode == 0 and output_file.exists() and output_file.stat().st_size > 0:
                        self.log(f"  ✓ 成功 ({output_file.stat().st_size} 字节)")
                        success_count += 1
                    else:
                        self.log(f"  ✗ 失败")
                        fail_count += 1

                except Exception as e:
                    self.log(f"  ✗ 错误: {str(e)}")
                    fail_count += 1

            # 显示总结
            self.log("\n" + "=" * 80)
            self.log("\n处理完成!")
            self.log(f"  总计: {len(zip_files)} 个文件")
            self.log(f"  成功: {success_count} 个文件")
            self.log(f"  失败: {fail_count} 个文件")
            self.log(f"\n解密文件保存在: {output_path}")

            messagebox.showinfo("完成",
                              f"处理完成!\n\n总计: {len(zip_files)} 个文件\n成功: {success_count} 个\n失败: {fail_count} 个\n\n解密文件保存在:\n{output_path}")

        except Exception as e:
            self.log(f"\n发生错误: {str(e)}")
            messagebox.showerror("错误", f"发生错误:\n{str(e)}")
        finally:
            self.start_button.config(state="normal")


def main():
    """主函数"""
    root = Tk()
    app = DecryptGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

## 结束 program #1 / End of program #1


# 下面是另一段用claude code生成的一段python代码，功能是提取一个文件夹下的包含特定文件名的xlsx文件里面固定单元格的内容，把这部分内容写入到一个xlsx的文件指定的位置 / Below is another python program created by claude code , the function is to extract the content of fixed cells from xlsx files containing a specific filename in a folder, and write this content to a specified location in a target xlsx file.
## 以下开始 program #2 / here begins program #2
import os
import glob
from openpyxl import load_workbook
import xlrd
import xlwt
from xlutils.copy import copy as xl_copy

# 源文件夹路径
source_folder = r"D:\LiptsDecryptTool\MSA-data\MC38290"
# 目标文件路径
target_file = r"D:\LiptsDecryptTool\MSA-data\MC38290\AIN_LVS_VDD12.xlsx"

# 查找所有包含CY089的xlsx文件（排除临时文件）
xlsx_files = []
for file in glob.glob(os.path.join(source_folder, "*CY089*.xlsx")):
    if not os.path.basename(file).startswith("~$"):
        xlsx_files.append(file)

print(f"找到 {len(xlsx_files)} 个文件")

# 按文件修改时间排序
xlsx_files.sort(key=lambda x: os.path.getmtime(x))

# 提取E26单元格的值
values = []
for file in xlsx_files:
    try:
        wb = load_workbook(file, data_only=True)
        # 尝试获取第一个工作表
        ws = wb.active
        # E26单元格
        cell_value = ws['E26'].value
        values.append(cell_value if cell_value is not None else "")
        print(f"{os.path.basename(file)}: {cell_value}")
        wb.close()
    except Exception as e:
        print(f"读取文件 {file} 时出错: {e}")
        values.append("")

print(f"\n提取到 {len(values)} 个值")

# 读取目标xlsx文件并写入G列
try:
    # 读取现有的xlsx文件
    wb = load_workbook(target_file)
    # 获取第2个工作表（sheet2）
    if len(wb.sheetnames) >= 2:
        ws = wb[wb.sheetnames[1]]
    else:
        print("警告：目标文件中没有sheet2，写入到第一个工作表")
        ws = wb.active

    # 写入G列，从第1行开始
    for i, value in enumerate(values):
        ws[f'G{i+1}'] = value  # G1, G2, G3...

    # 保存文件
    wb.save(target_file)
    wb.close()
    print(f"\n成功写入 {len(values)} 个值到 {target_file} 的sheet2的G列")

except Exception as e:
    print(f"写入目标文件时出错: {e}")

## 结束 program #2 / End of program #2

# 下面是最近整理的一些心得，讲如何规划一条产线和对应的工艺;/ Below is my own understanding of how to design a production line #
## 以下开始 program #3 / here begins program #3
如何设计一条产线？
步骤一：设定产线的核心指标：交付指标：1.年产能（20W或者40W），2.OEE指标（80%或85%），财务指标：1.成本，2.制造费用；质量指标：1.报废率；2.；
步骤二：有了产线的设计目标，下一步结合产品的实际开发情况，了解产品，拉通研发部门做DFMA（Design for Manufacture and Assembly），考虑产品设计是否适合导入量产；是否有优化的地方；DFMA可以伴随着产品的设计过程，同步迭代多轮；
步骤三：DFMA做完后，基本知道了产线需要做成什么样的，产品如何一步一步生产出来；下一步需要做LLD（Lean Line Design），拆解产品有多少打螺丝工站，多少焊接工站，多少测试工站，多少涂胶工站；需要规划这些工站的顺序，节拍，拆解各个工站的动作，评估各个工站的构成；其中有些工站会是难度大一点的工站，符合NUDD（New/Unique/Difficult/Different）；针对这类NUDD的工站，需要专门评估难度及风险；
步骤四：有了LLD，基本知道了整条产线的构成，多少个站位，节拍多少，人员多少，自动化率多少，下一步就可以写招标的技术文件了（SOR/SPEC）。技术文件需要尽量提到清楚，比如规定仪器的品牌及型号，用法用量，工站特殊要求，整条线体的特殊要求；需要特殊说明一点，厂房的电源/气源/网络/地坪也需要同步考虑。这类涉及到厂房资源往往是属于厂房土建的工作，但是也有可能前期规划厂房并未充分考虑，需要在产线规划的时候把边界沟通清楚；如厂房只提供整体供电，那么从变压器接电到具体的产线或者设备，这部分工作认为属于二次配（配电，配备网，配气）也需要考虑是否在产线招标的时候一并提出来；
步骤五：明确产线/设备的验收标准，这部分标准需要与研发/工厂/供应商/工艺都提前共识；验收标准涉及到对产线/设备要求的高低，要求越高，设备的制作精密程度越高，调试成本越高；设备产线验收：通常按照3-3-3-1的比例，30%预付款，供应商中标之后付款，30%预验收款，预验收通过后付款，30%终验收款，终验收通过后付款，10%质保尾款，质保期使用无问题后支付；预验收标准包括：设备清单与实物一致性检查，工具/工装清单与实物一致性检查，参考件清单与实物一致性检查，检具清单与实物一致性检查，能力验证(CmK/MSA/Cg/Cgk) ,防错验证,节拍验证,设备精度检查 (基础精度、设备对中),设备空循环 ,线体振动检查,红丹粉验证，产品通过性验证，实物质量检查，设备参数及限值的检查与优化，返工策略验证，旁路功能验证，连续生产4小时无因设备造成的质量缺陷，连续生产4小时设备停机时间小于15分钟，设备其它技术要求检查，MES需求检查，安全及人机工程检查，技术资料和图纸，操作培训，编程培训，维修培训，
终验收标准同上，有些细节可以改动，选择适用于预验收或者终验收；
步骤六：产线要求有了，验收标准有了，采购会进行招标定点；整体分为商务标+技术标，通常对供应商会有一个评分；一般评分分为两种标准，1.技术入围，最低价中标；2.技术评分占一定权重，商务评分占一定权重；最终综合得分高的供应商会拿到标的；
步骤七：招标定点后，与供应商进行技术对接，需要签订技术协议；技术协议可能会与当初的设备/产线SOR有偏差，需要有专门说明，留下文字记录以备后续审计；
步骤八：制订设备制作及调试计划，核对物料交期及品牌清单，定期开会检查设备制作的情况；沟通协调设备制作过程中的问题；
步骤九：按照标准，对设备进行预验收，组织研发/工厂/设备，对设备进行预验收，会用到前期协商好的标准，由各方签字认可；偏差项记入问题清单，后续跟进解决；
步骤十：预验收完成后，协商设备进工厂，接电接气接网，协商物料及调试；制订计划，
