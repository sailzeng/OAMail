from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout


if __name__ == '__main__':
    app = QApplication([])
    window = QWidget()
    layout = QVBoxLayout()
    layout.addWidget(QPushButton('Top'))
    layout.addWidget(QPushButton('Bottom'))
    window.setLayout(layout)
    window.show()
    app.exec()