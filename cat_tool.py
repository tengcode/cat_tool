import sys
from PySide6 import QtWidgets
from script import cat_tool

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    my_window = cat_tool.MyWidget()
    style_file = './script/tool.qss'
    # style_sheet = QSSLoader.read_qss_file(style_file)
    with open(style_file, 'r', encoding='UTF-8') as file:
        my_window.setStyleSheet(file.read())

    my_window.resize(700, 500)
    my_window.show()
    sys.exit(app.exec())
