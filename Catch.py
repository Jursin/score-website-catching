import sys
import os
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QCheckBox, QPushButton, QTextEdit, QFileDialog, QGroupBox, QFormLayout, QScrollArea
from PyQt5.QtCore import Qt
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.selected_files = []  # 用于存储选择的文件路径
        self.initUI()

    def initUI(self):
        # 主布局
        main_layout = QVBoxLayout()

        # 输入文件路径选择
        input_layout = QHBoxLayout()
        self.input_file_edit = QLineEdit(self)
        input_button = QPushButton("选择输入文件夹", self)
        input_button.clicked.connect(self.select_input_folder)  # 选择文件夹
        input_layout.addWidget(QLabel("输入文件夹:"))
        input_layout.addWidget(self.input_file_edit)
        input_layout.addWidget(input_button)

        # 创建文件显示区域
        self.file_list_groupbox = QGroupBox("Excel 文件列表")
        self.file_list_layout = QFormLayout()
        self.file_list_groupbox.setLayout(self.file_list_layout)

        # 可滚动的区域，并限制其大小
        scroll_area = QScrollArea()
        scroll_area.setWidget(self.file_list_groupbox)
        scroll_area.setWidgetResizable(True)  # 让QScrollArea根据内容自动调整大小
        scroll_area.setFixedHeight(150)  # 限制显示区域的高度
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)  # 禁用水平滚动条

        # 输入文件夹选择
        main_layout.addLayout(input_layout)
        main_layout.addWidget(scroll_area)  # 将文件列表滚动区域嵌入布局中

        # 输出文件路径
        output_layout = QHBoxLayout()
        self.output_file_edit = QLineEdit(self)
        output_button = QPushButton("选择输出路径", self)
        output_button.clicked.connect(self.select_output_path)
        output_layout.addWidget(QLabel("输出文件路径:"))
        output_layout.addWidget(self.output_file_edit)
        output_layout.addWidget(output_button)

        # 自动更改文件名勾选框
        self.auto_name_checkbox = QCheckBox("自动更改文件名", self)
        self.auto_name_checkbox.toggled.connect(self.auto_generate_filename)

        # 手动修改文件名文本框
        self.manual_filename_edit = QLineEdit(self)
        self.manual_filename_edit.setPlaceholderText("手动输入文件名（如果没有勾选自动更改）")

        # URL 输入框
        self.url_edit = QLineEdit(self)
        self.url_edit.setPlaceholderText("请输入网址")

        # 运行按钮
        run_button = QPushButton("运行", self)
        run_button.clicked.connect(self.run_process)

        # 控制台输出
        self.console_output = QTextEdit(self)
        self.console_output.setReadOnly(True)

        # 添加布局
        main_layout.addLayout(output_layout)
        main_layout.addWidget(self.manual_filename_edit)
        main_layout.addWidget(self.auto_name_checkbox)
        main_layout.addWidget(self.url_edit)
        main_layout.addWidget(run_button)
        main_layout.addWidget(self.console_output)

        # 设置窗口属性
        self.setLayout(main_layout)
        self.setWindowTitle('数据抓取工具')
        self.setGeometry(100, 100, 600, 500)
        self.show()

    def select_input_folder(self):
        folder_name = QFileDialog.getExistingDirectory(self, "选择输入文件夹")
        if folder_name:
            self.input_file_edit.setText(folder_name)
            self.load_files_in_folder(folder_name)

    def load_files_in_folder(self, folder_name):
        # 清空当前显示的文件列表
        for i in reversed(range(self.file_list_layout.count())):
            widget = self.file_list_layout.itemAt(i).widget()
            if widget is not None:
                widget.deleteLater()

        # 获取文件夹中的所有 Excel 文件
        excel_files = [f for f in os.listdir(folder_name) if f.endswith(('.xlsx', '.xls'))]
        self.checkboxes = []  # 用于存储所有勾选框

        for file in excel_files:
            checkbox = QCheckBox(file)
            self.checkboxes.append((checkbox, os.path.join(folder_name, file)))  # 存储勾选框和文件路径的元组
            self.file_list_layout.addRow(checkbox)  # 添加到文件列表显示区域

    def select_output_path(self):
        folder_name = QFileDialog.getExistingDirectory(self, "选择输出文件夹")
        if folder_name:
            self.output_file_edit.setText(folder_name)

    def auto_generate_filename(self):
        if self.auto_name_checkbox.isChecked():
            self.console_output.append("自动生成文件名已启用。")
            self.manual_filename_edit.setDisabled(True)  # 禁用手动修改文件名框
        else:
            self.manual_filename_edit.setDisabled(False)  # 恢复手动修改文件名框

    def run_process(self):
        selected_files = [file for checkbox, file in self.checkboxes if checkbox.isChecked()]
        if not selected_files:
            self.console_output.append("没有选择任何文件.")
            return

        output_folder = self.output_file_edit.text()  # 获取用户选择的输出文件夹路径

        # 如果没有勾选自动更改文件名，则使用手动输入的文件名
        if not self.auto_name_checkbox.isChecked():
            output_file = os.path.join(output_folder, self.manual_filename_edit.text())
        else:
            output_file = output_folder  # 默认输出到用户选择的文件夹

        # 显示控制台输出
        self.console_output.append("开始处理...")

        # 初始化 WebDriver
        driver = self.init_driver()
        if not driver:
            return

        # 批量处理选中的文件
        for input_file in selected_files:
            self.console_output.append(f"正在处理: {input_file}")
            self.process_data(input_file, output_file, driver)

        driver.quit()

    def init_driver(self):
        try:
            driver = webdriver.Edge()
            return driver
        except Exception as e:
            self.console_output.append(f"WebDriver 初始化失败: {e}")
            return None

    def process_data(self, input_file, output_file, driver):
        # 读取 Excel 文件
        df = pd.read_excel(input_file)
        self.console_output.append(f"读取输入文件: {input_file}")

        # 加载页面
        if not self.load_page(driver, self.url_edit.text()):
            return

        # 存储抓取到的数据
        output_data = []

        # 循环读取每一行的数据
        for index, row in df.iterrows():
            data1 = row.iloc[1]  # 从 Excel 表格中读取第二列的值
            data2 = str(row.iloc[2])[-4:]  # 从 Excel 表格中读取第三列的后四位

            try:
                # 使用显式等待定位输入框，超时设为3秒
                input1 = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME, 's_xingming')))
                input2 = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME, 's_chaxunma')))

                # 清除输入框中的内容
                input1.clear()
                input2.clear()

                # 将数据输入到输入框
                input1.send_keys(str(data1))
                input2.send_keys(str(data2))

                # 定位查询按钮并点击
                submit_button = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"查询")]')))
                submit_button.click()

                # 等待页面加载新的数据
                WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CLASS_NAME, 'right_cell')))

                # 抓取输出并保存
                elements = driver.find_elements(By.CLASS_NAME, 'right_cell')
                output_values = [element.text for element in elements]
                output_data.append([data1, data2] + output_values)

            except Exception as e:
                self.console_output.append(f"处理第 {index + 1} 行数据时出错: {e}")
                if not self.load_page(driver, self.url_edit.text()):  # 错误时重新加载页面
                    return
                continue

            driver.back()  # 返回上一页

        # 生成文件名并保存数据
        if output_data:
            try:
                new_filename = str(output_data[0][2]) + "班.xlsx"  # 使用抓取的第二列的值作为文件名
                output_folder = os.path.dirname(output_file)
                if not os.path.exists(output_folder):
                    os.makedirs(output_folder)  # 创建文件夹
                output_file = os.path.join(output_folder, new_filename)  # 将文件名拼接到输出路径中
                self.console_output.append(f"新的输出文件名: {output_file}")

                # 保存输出数据
                output_df = pd.DataFrame(output_data)
                output_df.to_excel(output_file, index=False)
                self.console_output.append(f"输出文件已保存: {output_file}")
            except PermissionError as e:
                self.console_output.append(f"保存文件失败: {e}")
                self.console_output.append("请检查目标文件夹是否存在，并确保程序有写入权限。")
            except Exception as e:
                self.console_output.append(f"保存文件时出错: {e}")

    def load_page(self, driver, url):
        try:
            driver.get(url)
            return True
        except Exception as e:
            self.console_output.append(f"页面加载失败: {e}")
            return False

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())