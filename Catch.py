import sys
import os
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QCheckBox, QPushButton, QTextEdit, QFileDialog, QGroupBox, QFormLayout, QScrollArea
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class SeleniumWorker(QThread):
    update_console = pyqtSignal(str)
    
    def __init__(self, input_files, url, output_path, auto_name, manual_filename):
        super().__init__()
        self.input_files = input_files
        self.url = url
        self.output_path = output_path
        self.auto_name = auto_name
        self.manual_filename = manual_filename

    def run(self):
        output_data = []
        driver = None
        try:
            driver = webdriver.Edge(executable_path='path/to/your/edgedriver')  # 更新为 WebDriver 的实际路径
            driver.get(self.url)
        except Exception as e:
            self.update_console.emit(f"WebDriver 错误: {e}")
            return

        for input_file in self.input_files:
            self.update_console.emit(f"处理文件: {input_file}")
            df = pd.read_excel(input_file)
            for index, row in df.iterrows():
                try:
                    data1 = row[1]
                    data2 = str(row[2])[-4:]
                    input1 = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME, 's_xingming')))
                    input2 = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME, 's_chaxunma')))
                    input1.clear()
                    input2.clear()
                    input1.send_keys(str(data1))
                    input2.send_keys(str(data2))
                    submit_button = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"查询")]')))
                    submit_button.click()
                    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CLASS_NAME, 'right_cell')))
                    elements = driver.find_elements(By.CLASS_NAME, 'right_cell')
                    output_values = [element.text for element in elements]
                    output_data.append([data1, data2] + output_values)
                except Exception as e:
                    self.update_console.emit(f"Error processing row {index}: {e}")
                    driver.get(self.url)
                    continue
                driver.get(self.url)

        if output_data:
            filename = self.manual_filename if not self.auto_name else str(output_data[0][2]) + "班.xlsx"
            output_file = os.path.join(self.output_path, filename)
            pd.DataFrame(output_data).to_excel(output_file, index=False)
            self.update_console.emit(f"文件已保存: {output_file}")

        if driver:
            driver.quit()

class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.selected_files = []
        self.initUI()

    def initUI(self):
        font = QFont()
        font.setPointSize(12)
        self.setFont(font)
        
        main_layout = QVBoxLayout()
        main_layout.setSpacing(15)

        # 输入文件夹选择布局
        input_layout = QHBoxLayout()
        self.input_file_edit = QLineEdit(self)
        self.input_file_edit.setFixedHeight(30)
        input_button = QPushButton("选择输入文件夹", self)
        input_button.setFixedSize(150, 40)
        input_button.clicked.connect(self.select_input_folder)
        input_layout.addWidget(QLabel("输入文件夹:"))
        input_layout.addWidget(self.input_file_edit)
        input_layout.addWidget(input_button)

        # Excel 文件列表布局
        self.file_list_groupbox = QGroupBox("Excel 文件列表")
        self.file_list_layout = QFormLayout()
        self.file_list_groupbox.setLayout(self.file_list_layout)

        scroll_area = QScrollArea()
        scroll_area.setWidget(self.file_list_groupbox)
        scroll_area.setWidgetResizable(True)
        scroll_area.setFixedHeight(200)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        main_layout.addLayout(input_layout)
        main_layout.addWidget(scroll_area)

        # 输出文件路径布局
        output_layout = QHBoxLayout()
        self.output_file_edit = QLineEdit(self)
        self.output_file_edit.setFixedHeight(30)
        output_button = QPushButton("选择输出路径", self)
        output_button.setFixedSize(150, 40)
        output_button.clicked.connect(self.select_output_path)
        output_layout.addWidget(QLabel("输出文件路径:"))
        output_layout.addWidget(self.output_file_edit)
        output_layout.addWidget(output_button)

        # 自动生成文件名复选框和手动输入文件名框布局
        filename_layout = QHBoxLayout()
        
        self.auto_name_checkbox = QCheckBox("自动更改文件名", self)
        self.auto_name_checkbox.setFixedHeight(30)
        self.auto_name_checkbox.toggled.connect(self.auto_generate_filename)
        
        self.manual_filename_edit = QLineEdit(self)
        self.manual_filename_edit.setPlaceholderText("手动输入文件名")
        self.manual_filename_edit.setFixedHeight(30)
        
        filename_layout.addWidget(self.auto_name_checkbox)
        filename_layout.addWidget(self.manual_filename_edit)

        # 输入网址和运行按钮布局
        url_layout = QHBoxLayout()
        self.url_edit = QLineEdit(self)
        self.url_edit.setPlaceholderText("请输入网址")
        self.url_edit.setFixedHeight(30)

        run_button = QPushButton("运行", self)
        run_button.setFixedSize(100, 40)
        run_button.clicked.connect(self.run_process)

        url_layout.addWidget(self.url_edit)
        url_layout.addWidget(run_button)

        # 控制台输出区域
        self.console_output = QTextEdit(self)
        self.console_output.setReadOnly(True)
        self.console_output.setFixedHeight(200)

        # 添加布局到主布局
        main_layout.addLayout(output_layout)
        main_layout.addLayout(filename_layout)
        main_layout.addLayout(url_layout)
        main_layout.addWidget(self.console_output)

        # 窗口设置
        self.setLayout(main_layout)
        self.setWindowTitle('数据抓取工具')
        self.setGeometry(100, 100, 800, 600)
        self.show()

    def select_input_folder(self):
        folder_name = QFileDialog.getExistingDirectory(self, "选择输入文件夹")
        if folder_name:
            self.input_file_edit.setText(folder_name)
            self.load_files_in_folder(folder_name)

    def load_files_in_folder(self, folder_name):
        for i in reversed(range(self.file_list_layout.count())):
            widget = self.file_list_layout.itemAt(i).widget()
            if widget is not None:
                widget.deleteLater()

        excel_files = [f for f in os.listdir(folder_name) if f.endswith(('.xlsx', '.xls'))]
        self.checkboxes = []
        for file in excel_files:
            checkbox = QCheckBox(file)
            self.checkboxes.append((checkbox, os.path.join(folder_name, file)))
            self.file_list_layout.addRow(checkbox)

    def select_output_path(self):
        folder_name = QFileDialog.getExistingDirectory(self, "选择输出文件夹")
        if folder_name:
            self.output_file_edit.setText(folder_name)

    def auto_generate_filename(self):
        if self.auto_name_checkbox.isChecked():
            self.console_output.append("自动生成文件名已启用。")
            self.manual_filename_edit.hide()
        else:
            self.manual_filename_edit.show()

    def run_process(self):
        selected_files = [file for checkbox, file in self.checkboxes if checkbox.isChecked()]
        if not selected_files:
            self.console_output.append("没有选择任何文件.")
            return

        url = self.url_edit.text()
        output_path = self.output_file_edit.text()
        auto_name = self.auto_name_checkbox.isChecked()
        manual_filename = self.manual_filename_edit.text()

        self.worker = SeleniumWorker(selected_files, url, output_path, auto_name, manual_filename)
        self.worker.update_console.connect(self.console_output.append)
        self.worker.start()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())
