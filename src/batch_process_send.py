# import office.execl_database
# import office.outlook_mail



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
        # EXCEL表的行列信息
        self.row_caption = 0
        self._row_start_read = 0
        self._column_sender = 0
        self._column_cc = 0
        self._column_body = 0
        self._column_attachment_list = []
        self._column_search_replace_list = []
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

        self._outlook = OutlookMail()
        self._excel = ExcelDataBase()


if __name__ == '__main__':
    pass
