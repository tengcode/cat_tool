import os
import sys
from PySide6 import QtCore, QtWidgets, QtGui
import cat_excel
# import xlsxwriter as xw


class MyWidget(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('合并文件')
        self.label_box = QtWidgets.QLabel("文件格式：xlsx，\n"
                                          "请把要合并的文件放到同一个文件夹，\n"
                                          "并输入路径：")
        self.text_box = QtWidgets.QLineEdit()
        self.run_btn = QtWidgets.QPushButton("确定")

        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.label_box)
        self.layout.addWidget(self.text_box)
        self.layout.addWidget(self.run_btn)

        self.run_btn.clicked.connect(self.show_message)

    @QtCore.Slot()
    def show_message(self):
        msg_box = QtWidgets.QMessageBox()
        path = self.text_box.text()
        if os.path.exists(path):
            cat_excel.cat_excel(file_path=path, new_file_name='数据归总')
            msg_box.information(self, "Note", f"正在合并文件，路径是： {path};\n"
                                              f"新文件为‘数据汇总’。")
            msg_box.setStandardButtons(QtWidgets.QMessageBox.Ok)
        else:
            msg_box.information(self, "Error", f"文件夹路径不正确，请检查。")


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)

    widget = MyWidget()
    widget.resize(500, 400)
    widget.show()
    sys.exit(app.exec())
