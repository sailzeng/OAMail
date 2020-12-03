#import execl_database
#import outlook_mail


class BatchProcessSend(object):
    #
    class SearchReplaceColumn(object):
        def __init__(self):
            self._search_column = 0
            self._replace_column = 0

    class BodyReplace(object):
        def __init__(self):
            self._search = ""
            self._replace = ""

    def __init__(self):
        #
        self._start_read_row = 0
        self._sender_column = 0
        self._cc_column = 0
        self._body_column = 0
        self._attachment_column_list = []
        self._search_replace_column_list = []

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
        self._subject_list =[]
        self._body_list = []
        self._attachments_list = []
        # Body 文本替换列表
        self._body_replace_list = []



if __name__ == '__main__':
    pass