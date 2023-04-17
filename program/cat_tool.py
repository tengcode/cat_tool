"""
Purpose    : cat excel tool
Programmer : Bruce Ma
Start date : 2023-04-11
"""

import os
import sys
from PySide6 import QtCore, QtWidgets, QtGui
import cat_excel
# import xlsxwriter as xw


class MyWidget(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('合并文件')
        self.type_list = {'合并到不同sheet': 'multi',
                          '合并到同一个sheet': 'single'}
        self.text_label = QtWidgets.QLabel("支持的文件格式：xlsx，\n"
                                           "请将文件放到同一个文件夹，并输入路径：")
        self.text_box = QtWidgets.QLineEdit()
        self.name_label = QtWidgets.QLabel("请输入新文件名并避免重复(默认名称为'数据汇总'):")
        self.name_box = QtWidgets.QLineEdit()
        self.type_box = QtWidgets.QComboBox()
        self.type_box.addItem('合并到不同sheet')
        self.type_box.addItem('合并到同一个sheet')
        self.run_btn = QtWidgets.QPushButton("确定")
        self.layout = QtWidgets.QVBoxLayout(self)
        self.set_layout()

    def set_layout(self):
        self.layout.addWidget(self.text_label)
        self.layout.addWidget(self.text_box)
        self.layout.addWidget(self.name_label)
        self.layout.addWidget(self.name_box)

        self.layout.addWidget(self.type_box)
        self.layout.addWidget(self.run_btn)

        self.run_btn.clicked.connect(self.start_cat)

    @QtCore.Slot()
    def start_cat(self):
        msg_box = QtWidgets.QMessageBox()
        path = self.text_box.text()
        name = self.name_box.text()
        if name == '':
            name = '数据汇总'
        cat_type = self.type_list[self.type_box.currentText()]
        if os.path.exists(path):
            msg_box.information(self,
                                "Note",
                                f"正在合并文件，路径是： {path};\n"
                                f"新文件为'{name}'。")
            cat_excel.cat_excel(file_path=path,
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


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)

    widget = MyWidget()
    widget.resize(400, 300)
    widget.show()
    sys.exit(app.exec())
