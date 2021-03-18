import office.outlook_mail as outlook
import office.excel_database as excel


class BatchProcessSend(object):
    """
    批处理进行发送处理
    """

    class SearchReplaceColumn(object):
        """
        搜索替换的列号
        """

        def __init__(self):
            self._column_search = 0
            self._column_replace = 0

    class BodyReplaceStr(object):
        """
        搜索替换的字符串
        """

        def __init__(self):
            self._search = ""
            self._replace = ""

    def __init__(self):
        # EXCEL 表的信息,根据EXCEL得到
        self._sheets_name = []
        self._sheet_row_start = 0
        self._sheet_column_start = 0
        self._sheet_row_end = 0
        self._sheet_column_count = 0
        self._sheet_data = []

        # 读取EXCEL的配置信息
        self._mail_row_caption = 0
        self._mail_row_start = 0
        self._mail_row_end = 0
        self._mail_count = 0

        self._column_sender = 0
        self._column_to = 1
        self._column_cc = 2
        self._column_subject = 3
        self._column_body = 4
        self._column_attachment_list = []
        self._column_search_replace_list = []

        # 默认邮件信息
        self._default_mail = outlook.SendMailInfo()

        # 列表，如果每人的信息不一样
        self._send_mail_list = []

        self._outlook = outlook.OutlookMail()
        self._excel = excel.ExcelDataBase()
        return

    def config_column(self,
                      column_sender: int = 0,
                      column_to: int = 1,
                      column_cc: int = 2,
                      column_subject: int = 3,
                      column_body: int = 4,
                      column_attachment_list: list = None,
                      column_search_list: list = None,
                      column_replace_list: list = None):
        """
        配置读取列信息
        :param column_sender:
        :param column_to:
        :param column_cc:
        :param column_body:
        :param column_attachment_list:
        :param column_search_list:
        :param column_replace_list:
        :return:
        """
        self._column_sender = column_sender
        self._column_to = column_to
        self._column_cc = column_cc
        self._column_body = column_body
        for column_attachment in column_attachment_list:
            self._column_attachment_list.append(column_attachment)
        len_s = len(column_search_list)
        len_r = len(column_replace_list)
        if len_s != len_r:
            return False
        for column_search, column_replace in zip(column_search_list, column_replace_list):
            zip_column = BatchProcessSend.SearchReplaceColumn()
            zip_column._column_search = column_search
            zip_column._column_replace = column_replace
            self._column_search_replace_list.append(zip_column)
        return True

    def config_default(self,
                       sender=None,
                       to=None,
                       cc=None,
                       subject=None,
                       body=None,
                       attachment_list=None):
        self._default_mail.sender = sender
        self._default_mail.to = to
        self._default_mail.cc = cc
        self._default_mail.subject = subject
        self._default_mail.body = body
        self._default_mail.attachment_list = attachment_list
        if sender is not None:
            self._column_sender = 0
        if to is not None:
            self._column_to = 0
        if cc is not None:
            self._column_cc = 0
        if subject is not None:
            self._column_subject = 0
        if body is not None:
            self._column_body = 0
        if attachment_list is not None:
            self._column_attachment_list = 0
        return True

    def open_excel(self, xls_file: str) -> bool:
        ret = self._excel.start()
        if not ret:
            return False
        ret = self._excel.open_book(xls_file, False)
        if not ret:
            return False
        self._sheets_name = self._excel.sheets_name()

        return True

    def load_sheet(self, sheet_name: str) -> bool:
        ret = self._excel.load_sheet_byname(sheet_name)
        if not ret:
            return False
        # 读取EXCEL Sheet的信息
        (self._sheet_row_start, self._sheet_column_start, self._sheet_row_end, self._sheet_column_count,
         self._sheet_data) = self._excel.used_range_data()

        self._mail_row_caption = self._sheet_row_start
        self._mail_row_start = self._sheet_row_start + 1
        self._mail_row_end = self._sheet_row_end
        self._mail_count = self._mail_row_end - self._mail_row_end + 1

        if self._column_sender != 0:
            self._column_sender += self._sheet_column_start
        if self._column_to != 0:
            self._column_to += self._sheet_column_start
        if self._column_cc != 0:
            self._column_cc += self._sheet_column_start
        if self._column_subject != 0:
            self._column_subject += self._sheet_column_start
        if self._column_body != 0:
            self._column_body += self._sheet_column_start
        return True

    def generate_run_para(self):
        self._mail_count = self._mail_row_end - self._mail_row_end + 1
        return

    def check_run_para(self) -> bool:
        if self._mail_row_start <= 0 or self._mail_row_end <= 0:
            return False
        if self._column_sender > 0 or self._column_to:
            return False
        return True

    def batch_send_list(self) -> bool:
        self.generate_run_para()
        if not self.check_run_para():
            return False
        i = 0
        while i < self._mail_count:
            new_mail = outlook.SendMailInfo()
            if 0 != self._column_sender:
                new_mail.sender = self._sheet_data[i + self._mail_row_start][self._column_sender]
            else:
                new_mail.sender = self._default_mail.sender
            if 0 != self._column_to:
                new_mail.to = self._sheet_data[i + self._mail_row_start][self._column_to]
            else:
                new_mail.to = self._default_mail.to
            if 0 != self._column_cc:
                new_mail.cc = self._sheet_data[i + self._mail_row_start][self._column_cc]
            else:
                new_mail.cc = self._default_mail.cc
            if 0 != self._column_subject:
                new_mail.subject = self._sheet_data[i + self._mail_row_start][self._column_subject]
            else:
                new_mail.subject = self._default_mail.subject
            if 0 != self._column_body:
                new_mail.body = self._sheet_data[i + self._mail_row_start][self._column_body]
            else:
                new_mail.body = self._default_mail.body

            if not self._column_attachment_list:
                for column_attachment in self._column_attachment_list:
                    new_mail.attachment_list.append(column_attachment)

            self._outlook.create_sendmail()
            self._outlook.set_send_account(new_mail.sender)
            self._outlook.set_sendmail(new_mail.to, new_mail.cc, new_mail.subject, new_mail.body)
            self._outlook.send_mail()

        return True


if __name__ == '__main__':
    batch_send = BatchProcessSend()
    batch_send.config_column(0,1,2,3,4)
    batch_send.config_default("sailzeng@tencent.com")
    batch_send.
    pass
