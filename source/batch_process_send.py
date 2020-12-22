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
        搜索替换的语句
        """
        def __init__(self):
            self._search = ""
            self._replace = ""

    def __init__(self):
        # EXCEL 表的信息
        self._sheets_name = []
        self._sheet_row_start = 0
        self._sheet_column_start = 0
        self._sheet_row_count = 0
        self._sheet_column_count = 0

        # 读取EXCEL的配置信息
        self._row_caption = 0
        self._row_start_read = 0
        self._row_count = 0

        self._column_sender = 0
        self._column_to = 0
        self._column_cc = 0
        self._column_body = 0
        self._column_attachment_list = []
        self._column_search_replace_list = []

        self._sendmail_count = 0

        # 发送单人
        self._sender = ""
        self._cc = ""
        self._subject = ""
        self._body = ""
        self._attachments = []

        # 列表，如果每人的信息不一样
        self._sender_list = []
        self._to_list = []
        self._cc_list = []
        self._subject_list = []
        self._body_list = []
        self._attachments_list = []
        # Body 文本替换列表
        self._body_replace_list = []

        self._outlook = outlook.OutlookMail()
        self._excel = excel.ExcelDataBase()
        return

    def config_column(self,
                      column_sender: int = 0,
                      column_cc: int = 0,
                      column_body: int = 0,
                      column_attachment_list: list = [],
                      column_search_list: list = [],
                      column_replace_list: list = []):
        """

        :param column_sender:
        :param column_cc:
        :param column_body:
        :param column_attachment_list:
        :param column_search_list:
        :param column_replace_list:
        :return:
        """
        self._column_sender = column_sender
        self._column_cc = column_cc
        self._column_body = column_body
        for column_attachment in column_attachment_list:
            self._column_attachment_list.append(column_attachment)
        len_s = len(column_search_list)
        len_r = len(column_replace_list)
        if len_s != len_r:
            if len_s > 0 and len_r == 0:
                # 没有填写替换数据
                return False
            if len_s == 0:
                #
                column_search_list =
        for column_search, column_replace in zip(column_search_list,column_replace_list):
            zip_column = BatchProcessSend.SearchReplaceColumn()
            zip_column._column_search = column_search
            zip_column._column_replace = column_replace
            self._column_attachment_list.append(zip_column)
        return True

    def open_excel(self, xls_file: str) -> bool:
        ret = self._excel.start()
        if not ret :
            return False
        ret = self._excel.open_book(xls_file, False)
        if not ret:
            return False
        self._sheets_name = self._excel.sheets_name()

        return True

    def load_sheet(self, sheet_name: str) -> bool:
        ret = self._excel.load_sheet(sheet_name)
        if not ret:
            return False
        self._sheet_row_start, self._sheet_column_start, self._sheet_row_count, self._sheet_column_count \
            = self._excel.used_range_coord()
        return True

    def check_config(self) -> bool:

        if self._sendmail_count > self._sheet_row_count - self._sheet_row_start + 1:
            return False
        if self._row_start_read > self._sheet_row_start or self._row_caption > self._sheet_row_start:
            return False
        self._row_count = self._sheet_row_count - self._row_start_read
        if self._sendmail_count <= 0:
            self._sendmail_count = self._row_count
        if self._sendmail_count > self._row_count:
            self._sendmail_count = self._row_count

        return True

    def generate_send_list(self)-> bool:

        return True


if __name__ == '__main__':
    pass
