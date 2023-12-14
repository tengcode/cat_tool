"""
Purpose    : cat excel tool
Programmer : Bruce Ma
Start date : 2023-04-11
Update     : 2023-12-14
"""

import os
from PySide6 import QtCore, QtWidgets, QtGui
from PySide6.QtWidgets import (QMainWindow, QWidget, QLabel, QLineEdit,
                               QFileDialog, QComboBox, QPushButton, QVBoxLayout)
from script import cat


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('合并文件')
        self.centralWidget = QWidget()
        self.setCentralWidget(self.centralWidget)
        self.type_list = {'合并到不同sheet': 'multi',
                          '合并到同一个sheet': 'single'}
        self.text_label = QLabel("支持的文件格式：xlsx，sas7bdat, csv\n"
                                 "请将文件放到同一个文件夹，并选择文件夹路径：")
        self.folder_box = QLineEdit()
        self.file_check_btn = QPushButton('选择文件夹')
        self.name_label = QLabel("请输入新文件名并避免重复(默认名称为'数据汇总'):")
        self.name_box = QLineEdit()
        self.type_box = QComboBox()
        self.type_box.addItem('合并到不同sheet')
        self.type_box.addItem('合并到同一个sheet')
        self.run_btn = QPushButton("确定")
        self.set_layout()

    def set_layout(self):
        layout = QVBoxLayout(self)
        layout.addWidget(self.text_label)
        layout.addWidget(self.folder_box)
        layout.addWidget(self.file_check_btn)
        layout.addWidget(self.name_label)
        layout.addWidget(self.name_box)

        layout.addWidget(self.type_box)
        layout.addWidget(self.run_btn)

        self.centralWidget.setLayout(layout)
        self.run_btn.clicked.connect(self.start_cat)
        self.file_check_btn.clicked.connect(self.select_folder)

    @QtCore.Slot()
    def select_folder(self):
        file_path = QFileDialog.getExistingDirectory(self.centralWidget,
                                                     "选择存储路径", r'C:\Users\user\Desktop')
        self.folder_box.setText(file_path)


    @QtCore.Slot()
    def start_cat(self):
        msg_box = QtWidgets.QMessageBox()
        path = self.folder_box.text()
        name = self.name_box.text()
        if name == '':
            name = '数据汇总'
        cat_type = self.type_list[self.type_box.currentText()]
        if os.path.exists(path):
            msg_box.information(self,
                                "Note",
                                f"正在合并文件，路径是： {path};\n"
                                f"新文件为'{name}'。")
            cat.cat_data_file(file_path=path,
                              cat_type=cat_type,
                              new_file_name=name)
            msg_box.information(self,
                                "Note",
                                f"文件已合并完成，请查看！")
            msg_box.setStandardButtons(QtWidgets.QMessageBox.Ok)
        else:
            msg_box.information(self,
                                "Error",
                                f"文件夹路径不正确，请检查。")
        # ret = msg_box.exec()
