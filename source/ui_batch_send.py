
from PyQt5.QtWidgets import QApplication, QMainWindow, QStatusBar, QButtonGroup, QPushButton, QGridLayout, QLabel, \
    QRadioButton, QFileDialog, QMessageBox, QLineEdit, QTextEdit, QFrame, QWidget, QComboBox
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon

import batch_send as bps
import office.outlook_mail as outlook
from batch_send import RowNumConfig
import ctypes


class BatchSendMailWidget(QMainWindow):
    """
    批处理进行邮件发送处理处理的UI
    """
    _row_config: RowNumConfig

    def __init__(self):
        super().__init__()
        # UI处理的各种成员
        self.setCentralWidget(QWidget())
        self.setWindowIcon(QIcon('./icon/mail.png'))
        ctrl_line = []
        self._ctrl_matrix = []
        self._btn_group = []
        label = QLabel("发件人：")
        ctrl_line.append(label)
        radio1 = QRadioButton("相应EXCEL列号：")
        ctrl_line.append(radio1)
        edit = QLineEdit("1")
        ctrl_line.append(edit)
        radio2 = QRadioButton("使用填写发件人：")
        ctrl_line.append(radio2)
        edit = QLineEdit("")
        edit.setEnabled(False)
        ctrl_line.append(edit)
        group = QButtonGroup()
        group.addButton(radio1)
        group.addButton(radio2)
        radio1.clicked.connect(self.radio_clicked)
        radio2.clicked.connect(self.radio_clicked)
        group.setExclusive(True)
        radio1.setChecked(True)
        self._ctrl_matrix.append(ctrl_line)
        self._btn_group.append(group)
        ctrl_line = []
        label = QLabel("收件人：")
        ctrl_line.append(label)
        radio1 = QRadioButton("相应EXCEL列号：")
        ctrl_line.append(radio1)
        edit = QLineEdit("2")
        ctrl_line.append(edit)
        radio2 = QRadioButton("使用填写收件人：")
        ctrl_line.append(radio2)
        edit = QLineEdit("")
        edit.setEnabled(False)
        ctrl_line.append(edit)
        group = QButtonGroup()
        group.addButton(radio1)
        group.addButton(radio2)
        radio1.clicked.connect(self.radio_clicked)
        radio2.clicked.connect(self.radio_clicked)
        group.setExclusive(True)
        radio1.setChecked(True)
        self._ctrl_matrix.append(ctrl_line)
        self._btn_group.append(group)
        ctrl_line = []
        label = QLabel("抄送：")
        ctrl_line.append(label)
        radio1 = QRadioButton("相应EXCEL列号：")
        ctrl_line.append(radio1)
        edit = QLineEdit("3")
        ctrl_line.append(edit)
        radio2 = QRadioButton("使用填写抄送人：")
        ctrl_line.append(radio2)
        edit = QLineEdit("")
        edit.setEnabled(False)
        ctrl_line.append(edit)
        group = QButtonGroup()
        group.addButton(radio1)
        group.addButton(radio2)
        radio1.clicked.connect(self.radio_clicked)
        radio2.clicked.connect(self.radio_clicked)
        radio1.setChecked(True)
        self._ctrl_matrix.append(ctrl_line)
        self._btn_group.append(group)
        ctrl_line = []
        label = QLabel("邮件标题：")
        ctrl_line.append(label)
        radio1 = QRadioButton("相应EXCEL列号：")
        ctrl_line.append(radio1)
        edit = QLineEdit("4")
        ctrl_line.append(edit)
        radio2 = QRadioButton("使用填写标题：")
        ctrl_line.append(radio2)
        edit = QLineEdit("")
        edit.setEnabled(False)
        ctrl_line.append(edit)
        group = QButtonGroup()
        group.addButton(radio1)
        group.addButton(radio2)
        radio1.clicked.connect(self.radio_clicked)
        radio2.clicked.connect(self.radio_clicked)
        radio1.setChecked(True)
        self._ctrl_matrix.append(ctrl_line)
        self._btn_group.append(group)
        ctrl_line = []
        label = QLabel("邮件正文：")
        ctrl_line.append(label)
        radio1 = QRadioButton("相应EXCEL列号：")
        ctrl_line.append(radio1)
        edit = QLineEdit("5")
        ctrl_line.append(edit)
        radio2 = QRadioButton("使用填写正文：")
        ctrl_line.append(radio2)
        edit = QTextEdit("")
        edit.setEnabled(False)
        ctrl_line.append(edit)
        group = QButtonGroup()
        group.addButton(radio1)
        group.addButton(radio2)
        radio1.clicked.connect(self.radio_clicked)
        radio2.clicked.connect(self.radio_clicked)
        radio1.setChecked(True)
        self._ctrl_matrix.append(ctrl_line)
        self._btn_group.append(group)
        self._edit_open_xls = QLineEdit("")
        self._btn_open_xls = QPushButton("打开")
        self._btn_open_xls.clicked.connect(self.open_xls_clicked)
        self._combo_xls_sheet = QComboBox()
        self._combo_xls_sheet.currentIndexChanged.connect(self.combo_sheet_changed)
        self._edit_start_row = QLineEdit("2")
        self._edit_end_row = QLineEdit("5")
        self._btn_send = QPushButton("发送")
        self._btn_exit = QPushButton("退出")
        self._btn_send.clicked.connect(self.send_mail)
        self._btn_exit.clicked.connect(self.exit_app)
        self._status_bar = QStatusBar()
        self._status_bar.layout().setAlignment(Qt.AlignRight)
        # BPS批处理发送处理类
        self._batch_send = bps.BatchProcessSend()
        self._default_mail = outlook.SendMailInfo()
        self._row_config = bps.RowNumConfig()
        self._column_config = bps.ColumnNumConfig()

    def show(self):
        grid = QGridLayout()
        line = 0
        grid.setSpacing(10)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(2, 1)
        grid.setColumnStretch(3, 1)
        grid.setColumnStretch(4, 10)
        grid.setColumnStretch(5, 1)
        self.centralWidget().setLayout(grid)
        grid.addWidget(QLabel("打开EXCEL配置文件："), line, 0)
        grid.addWidget(self._edit_open_xls, line, 1, 1, 5)
        grid.addWidget(self._btn_open_xls, line, 6)
        line += 1
        grid.addWidget(QLabel("EXCEL对应的Sheet："), line, 0)
        grid.addWidget(self._combo_xls_sheet, line, 1, 1, 5)
        line += 1
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Plain)
        grid.addWidget(separator, line, 0, 1, 8)
        line += 1
        for x in self._ctrl_matrix:
            grid.addWidget(x[0], line, 0, 1, 1)
            grid.addWidget(x[1], line, 1, 1, 1)
            grid.addWidget(x[2], line, 2, 1, 1)
            grid.addWidget(x[3], line, 3, 1, 1)
            # 给最后一个TextEdit多留一点空间
            if line == 5:
                grid.addWidget(x[4], line, 4, 3, 5)
                line += 3
            else:
                grid.addWidget(x[4], line, 4, 1, 5)
                line += 1
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Plain)
        grid.addWidget(separator, line, 0, 1, 8)
        line += 1
        grid.addWidget(QLabel("EXCEL起始发送行："), line, 0)
        grid.addWidget(self._edit_start_row, line, 1)
        line += 1
        grid.addWidget(QLabel("EXCEL结束发送行："), line, 0)
        grid.addWidget(self._edit_end_row, line, 1)
        line += 1
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Plain)
        grid.addWidget(separator, line, 0, 1, 8)
        line += 1
        grid.addWidget(self._btn_send, line, 5)
        grid.addWidget(self._btn_exit, line, 6)
        self.move(300, 150)
        self.setWindowTitle('OA Send Mail')
        self.setStatusBar(self._status_bar)
        self._status_bar.showMessage("选择配置的EXCEL文件，进行批量发送")
        QMainWindow.show(self)

    def radio_clicked(self):
        radio = self.sender()
        for x in self._ctrl_matrix:
            if x[1] == radio:
                x[1].setChecked(True)
                x[2].setEnabled(True)
                x[4].setEnabled(False)
            elif x[3] == radio:
                x[3].setChecked(True)
                x[4].setEnabled(True)
                x[2].setEnabled(False)
        pass

    def open_xls_clicked(self):
        """
        打开文件对话框，
        """
        file_name, file_type = QFileDialog.getOpenFileName(
            self,
            "选取EXCEL文件",
            "./",
            "Excel Files (*.xls *.xlsx);;All Files (*)")
        if file_name == "":
            self._status_bar.showMessage("必须选择配置的EXCEL文件，才能进行批量发送")
            pass
        self._status_bar.showMessage("已经选择了EXCEL，请选择Sheet")
        self._edit_open_xls.setText(file_name)
        success_open = self._batch_send.open_excel(file_name)
        if not success_open:
            QMessageBox.information(self,
                                    "提示",
                                    "无法打开对应的EXCEL文件。",
                                    QMessageBox.Ok)
        self._combo_xls_sheet.clear()
        self._combo_xls_sheet.addItems(self._batch_send.sheets_name)
        pass

    def send_mail(self) -> bool:
        """
        发送邮件，
        """
        if self._ctrl_matrix[0][1].isChecked():
            self._column_config.sender = int(self._ctrl_matrix[0][2].text())
            if self._column_config.sender == 0:
                self._status_bar.showMessage("没有填写默认发送人时，发送人列号不能是0")
                return False
        else:
            self._column_config.sender = 0
            self._default_mail.sender = self._ctrl_matrix[0][4].text()
        if self._ctrl_matrix[1][1].isChecked():
            self._column_config.to = int(self._ctrl_matrix[1][2].text())
            if self._column_config.to == 0:
                self._status_bar.showMessage("没有填写默认接收人时，接收人列号不能是0")
                return False
            else:
                if self._column_config.sender == self._column_config.to:
                    self._status_bar.showMessage("不能出现相同的列号")
                    return False
        else:
            self._column_config.to = 0
            self._default_mail.to = self._ctrl_matrix[1][4].text()
        if self._ctrl_matrix[2][1].isChecked():
            self._column_config.cc = int(self._ctrl_matrix[2][2].text())
            if self._column_config.cc == 0:
                self._status_bar.showMessage("没有填写默认抄送人时，抄送人列号不能是0")
                return False
            else:
                if self._column_config.sender == self._column_config.to or \
                        self._column_config.to == self._column_config.cc:
                    self._status_bar.showMessage("不能出现相同的列号")
                    return False
        else:
            self._column_config.cc = 0
            self._default_mail.cc = self._ctrl_matrix[2][4].text()
        if self._ctrl_matrix[3][1].isChecked():
            self._column_config.subject = int(self._ctrl_matrix[3][2].text())
            if self._column_config.subject == 0:
                self._status_bar.showMessage("没有填写默认标题时，标题列号不能是0")
                return False
        else:
            self._column_config.subject = 0
            self._default_mail.subject = self._ctrl_matrix[3][4].text()
        if self._ctrl_matrix[4][1].isChecked():
            self._column_config.body = int(self._ctrl_matrix[4][2].text())
            if self._column_config.body == 0:
                self._status_bar.showMessage("没有填写默认内容时，内容列号不能是0")
                return False
            else:
                if self._column_config.sender == self._column_config.to or \
                        self._column_config.to == self._column_config.cc:
                    self._status_bar.showMessage("不能出现相同的列号")
                    return False
        else:
            self._column_config.body = 0
            self._default_mail.body = self._ctrl_matrix[4][4].text()

        self._row_config.send_start = int(self._edit_start_row.text())
        self._row_config.send_end = int(self._edit_end_row.text())
        if self._row_config.send_start >= self._row_config.send_end:
            self._status_bar.showMessage("起始，结束发送行错误")
            return False
        self._row_config.mail_count = self._row_config.send_end - self._row_config.send_start + 1
        cfg_success = self._batch_send.config(self._column_config,
                                              self._row_config,
                                              self._default_mail)
        if not cfg_success:
            self._status_bar.showMessage("配置存在问题，请检查")
            return False
        else:
            self._status_bar.showMessage("配置正确，准备发送{}封邮件".format(self._row_config.mail_count))
        i = 0
        j = self._row_config.send_start
        while j <= self._row_config.send_end:
            send_result = self._batch_send.send_one(j)
            if not send_result:
                self._status_bar.showMessage(str("发送第{}封,第{}行邮件失败").format(i + 1, j))
                return False
            j += 1
            i += 1
            self._status_bar.showMessage(str("发送第{}封,第{}行邮件成功").format(i + 1, j))
        self._status_bar.showMessage("全部发送完成")
        return True

    def exit_app(self):
        self._status_bar.showMessage("下次再见")
        QApplication.quit()

    def combo_sheet_changed(self):
        load_ok = self._batch_send.load_sheet_byname(self._combo_xls_sheet.currentText())
        if not load_ok:
            self._status_bar.showMessage("加载EXCEL对应的sheet失败，请检查表格")
            return
        out_info = "加载sheet[{}]成功.起始结束行列{}".format(
            self._combo_xls_sheet.currentText(),
            self._batch_send.cur_sheet_info())
        self._status_bar.showMessage(out_info)
        s_r, s_c, e_r, e_c = self._batch_send.cur_sheet_info()
        self._row_config.caption = s_r
        self._row_config.send_start = s_r + 1
        self._row_config.send_end = e_r
        self._row_config.mail_count = e_r - s_r
        self._edit_start_row.setText(str(self._row_config.send_start))
        self._edit_end_row.setText(str(self._row_config.send_end))

    def bps_start(self) -> bool:
        start_ok = self._batch_send.start()
        if not start_ok:
            QMessageBox.information(window,
                                    "提示",
                                    "OA Send Mail需要Outlook和Excel，否则无法使用。",
                                    QMessageBox.Ok)
            QApplication.quit()
        return True

    def closeEvent(self, event):
        """
        重写closeEvent方法，实现dialog窗体关闭时执行一些代码
        """
        self._batch_send.close()


if __name__ == '__main__':
    app = QApplication([])
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("OAMail")
    font = app.font()
    font.setPointSize(11)
    app.setFont(font)
    window = BatchSendMailWidget()
    ret = window.bps_start()
    window.show()
    app.exec()
