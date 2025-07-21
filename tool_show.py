import sys
import os
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, 
    QPushButton, QGroupBox, QTextEdit, QFileDialog, QRadioButton, QButtonGroup
)
from PyQt6.QtCore import Qt


class ConfigTableChecker(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("配置表检查工具-增强版")
        self.setup_ui()
        
    def setup_ui(self):
        # 主布局
        main_layout = QVBoxLayout()
        
        # 标题
        title_label = QLabel("配置表检查工具-增强版")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        main_layout.addWidget(title_label)
        
        # 表根目录部分
        dir_group = QGroupBox()
        dir_layout = QHBoxLayout()
        
        self.dir_label = QLabel("表根目录")
        self.dir_input = QLineEdit()
        self.dir_input.setPlaceholderText("例如: F:/script")
        self.browse_button = QPushButton("浏览")
        self.browse_button.clicked.connect(self.browse_directory)
        
        dir_layout.addWidget(self.dir_label)
        dir_layout.addWidget(self.dir_input)
        dir_layout.addWidget(self.browse_button)
        dir_group.setLayout(dir_layout)
        main_layout.addWidget(dir_group)
        
        # 检查按钮
        self.check_button = QPushButton("开始检查")
        self.check_button.clicked.connect(self.start_check)
        self.check_button.setStyleSheet("font-size: 14px; padding: 5px;")
        main_layout.addWidget(self.check_button, alignment=Qt.AlignmentFlag.AlignCenter)
        
        # 检查结果部分
        self.result_group = QGroupBox("检查结果")
        result_layout = QVBoxLayout()
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setPlaceholderText("文件检查结果将显示在这里...")
        result_layout.addWidget(self.result_text)
        self.result_group.setLayout(result_layout)
        main_layout.addWidget(self.result_group)
        
        # 规则选择部分（新增）
        self.rule_group = QGroupBox("请选择规则")
        rule_layout = QVBoxLayout()
        
        # 创建单选按钮组
        self.rule_btn_group = QButtonGroup(self)
        
        # 三个规则选项
        self.rule1 = QRadioButton("规则1: 检查ID是否漏填")
        self.rule2 = QRadioButton("规则2: 检查ID是否重复")
        self.rule3 = QRadioButton("规则3: 检查ID是否超出范围")
        
        # 添加到按钮组
        self.rule_btn_group.addButton(self.rule1, 1)
        self.rule_btn_group.addButton(self.rule2, 2)
        self.rule_btn_group.addButton(self.rule3, 3)
        
        rule_layout.addWidget(self.rule1)
        rule_layout.addWidget(self.rule2)
        rule_layout.addWidget(self.rule3)
        self.rule_group.setLayout(rule_layout)
        main_layout.addWidget(self.rule_group)
        
        # 执行规则按钮（新增）
        self.execute_button = QPushButton("执行规则")
        self.execute_button.clicked.connect(self.execute_rule)
        self.execute_button.setStyleSheet("font-size: 14px; padding: 5px;")
        self.execute_button.setEnabled(False)  # 初始不可用
        main_layout.addWidget(self.execute_button, alignment=Qt.AlignmentFlag.AlignCenter)
        
        # 规则执行结果输出框（新增）
        self.output_group = QGroupBox("规则执行结果")
        output_layout = QVBoxLayout()
        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)
        self.output_text.setPlaceholderText("规则执行结果将显示在这里...")
        output_layout.addWidget(self.output_text)
        self.output_group.setLayout(output_layout)
        main_layout.addWidget(self.output_group)
        
        self.setLayout(main_layout)
        
    def browse_directory(self):
        """浏览目录并设置到输入框中"""
        directory = QFileDialog.getExistingDirectory(self, "选择表根目录")
        if directory:
            self.dir_input.setText(directory)
    
    def find_excel_files(self, directory):
        """查找目录中的所有Excel文件"""
        excel_files = []
        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.lower().endswith(('.xlsx', '.xls')):
                    excel_files.append(os.path.join(root, file))
        return excel_files
    
    def start_check(self):
        """执行检查操作"""
        directory = self.dir_input.text()
        
        if not directory:
            self.result_text.setPlainText("错误：请先选择表根目录")
            return
        
        if not os.path.isdir(directory):
            self.result_text.setPlainText(f"错误：目录不存在\n{directory}")
            return
        
        # 查找Excel文件
        try:
            excel_files = self.find_excel_files(directory)
            result = f"搜索目录: {directory}\n"
            result += f"找到文件数: {len(excel_files)}\n"
            result += "文件列表:\n"
            for i, file in enumerate(excel_files):
                result += f"{i+1}. {file}\n"

            self.result_text.setPlainText(result)
            self.execute_button.setEnabled(True)  # 检查完成后启用执行按钮s
            self.excel_files = excel_files  # 保存文件列表供规则使用
            
        except Exception as e:
            self.result_text.setPlainText(f"检查过程中发生错误:\n{str(e)}")
            self.execute_button.setEnabled(False)
    
    def execute_rule(self):
        """执行选中的规则"""
        selected_rule = self.rule_btn_group.checkedId()
        
        if not hasattr(self, 'excel_files') or not self.excel_files:
            self.output_text.setPlainText("错误：请先执行文件检查")
            return
        
        if selected_rule == -1:
            self.output_text.setPlainText("错误：请先选择一个规则")
            return
        
        try:
            output = "=== 规则执行结果 ===\n"
            output += f"执行的规则: 规则{selected_rule}\n"
            output += f"检查文件数: {len(self.excel_files)}\n"
            
            # 根据选择的规则执行不同的逻辑
            if selected_rule == 1:
                output += self._execute_rule1()
            elif selected_rule == 2:
                output += self._execute_rule2()
            elif selected_rule == 3:
                output += self._execute_rule3()
                
            self.output_text.setPlainText(output)
            
        except Exception as e:
            self.output_text.setPlainText(f"规则执行过程中发生错误:\n{str(e)}")
    
    def _execute_rule1(self):
        # 获取当前的所有Excel文件路径
        directory = self.dir_input.text()
        excel_files_root = self.find_excel_files(directory)
        sheet_name = "WKshenshou"  # 假设要检查的工作表名
        output = ""
        empty_rows = []
        if not excel_files_root:
            return "没有找到任何Excel文件。\n"
        
        for file in excel_files_root:
            try:
                output += f"正在检查文件: {file}\n"
                df = pd.read_excel(file, sheet_name=sheet_name, header=5, usecols=[1])
                data = df.iloc[:, 0].tolist()
                for index, value in enumerate(data):
                    if pd.isna(value) or value == '':
                        empty_rows.append((file, index + 7))
                if empty_rows:
                    # 找到最后一个非空值的索引
                    last_non_empty = len(data) - 1
                    while last_non_empty >= 0 and (pd.isna(data[last_non_empty]) or str(data[last_non_empty]).strip() == ''):
                        last_non_empty -= 1
                    # 移除结尾连续空行 (行号 > last_non_empty + 7)
                    empty_rows = [row for row in empty_rows if row[1] <= (last_non_empty + 7)]
                    for file, row in empty_rows:
                        output += f"文件 {file} 第 {row} 行没有填写ID。\n"
                    empty_rows.clear()
            except Exception as e:
                output += f"文件 {file} 读取失败: {str(e)}\n"
            continue
        
        return output



    def _execute_rule2(self):
        pass
    
    def _execute_rule3(self):
        pass


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ConfigTableChecker()
    window.show()
    sys.exit(app.exec())