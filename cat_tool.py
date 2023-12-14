import sys
from PySide6 import QtWidgets
from script import cat_ui

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    my_window = cat_ui.MyWindow()
    style_file = './script/tool.qss'
    with open(style_file, 'r', encoding='UTF-8') as file:
        my_window.setStyleSheet(file.read())

    my_window.resize(700, 500)
    my_window.show()
    sys.exit(app.exec())
