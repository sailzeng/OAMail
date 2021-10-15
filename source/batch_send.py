import time
import copy
import office.outlook_mail as outlook
import office.excel_database as excel


class ColumnNumConfig(object):
    """
    列号配置
    """
    def __init__(self):
        # 发送者配置列，如果使用默认值，配置为0
        self.sender = 0
        # 接受者配置列，如果使用默认值，配置为0
        self.to = 1
        # 抄送配置列，如果使用默认值，配置为0
        self.cc = 2
        # 标题配置列，如果使用默认值，配置为0
        self.subject = 3
        # 邮件内容配置列， 如果使用默认值，配置为0
        self.body = 4
        # 附件配置列列表，（可以多个附件），如果没有配置为None
        self.attachment_list = []
        # 搜索/替换队列标签配置列列表，（可以多个），如果没有配置为None
        self.search_replace_list = []


class RowNumConfig(object):
    """
    行号配置
    """

    def __init__(self):
        # 标题行号
        self.caption = 0
        # 起始发送行
        self.send_start = 0
        # 结束发送行
        self.send_end = 0
        # 发送的总数
        self.mail_count = 0


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


class BatchProcessSend(object):
    """
    批处理进行邮件发送处理
    """
    _row_cfg: RowNumConfig
    _column_cfg: ColumnNumConfig
    _default_mail: outlook.SendMailInfo

    def __init__(self):
        # EXCEL 表的信息,根据EXCEL得到
        self._sheets_name = []
        self._sheet_row_start = 0
        self._sheet_column_start = 0
        self._sheet_row_end = 0
        self._sheet_column_end = 0
        self._sheet_data = []

        # 读取EXCEL的配置信息
        self._column_cfg = ColumnNumConfig()
        self._row_cfg = RowNumConfig()

        # 默认邮件信息
        self._default_mail = outlook.SendMailInfo()

        # 列表，如果每人的信息不一样
        self._send_mail_list = []

        self._outlook = None
        self._excel = None
        return

    def start(self):
        self._outlook = outlook.OutlookMail()
        self._excel = excel.ExcelDataBase()
        ret = self._outlook.start()
        if not ret:
            return False
        ret = self._excel.start()
        if not ret:
            return False
        return True

    def config(self,
               cfg_column: ColumnNumConfig,
               cfg_row: RowNumConfig,
               cfg_default: outlook.SendMailInfo):
        """
        配置读取列,读取行信息，默认值信息
        """
        self._column_cfg = copy.deepcopy(cfg_column)
        self._row_cfg = copy.deepcopy(cfg_row)
        self._row_cfg._mail_count = self._row_cfg.send_end - self._row_cfg.send_start + 1
        self._default_mail = copy.deepcopy(cfg_default)
        if cfg_default.sender != "":
            self._column_cfg.sender = 0
        if cfg_default.to != "":
            self._column_cfg.to = 0
        if cfg_default.cc != "":
            self._column_cfg.cc = 0
        if cfg_default.subject != "":
            self._column_cfg.subject = 0
        if cfg_default.body != "":
            self._column_cfg._body = 0
        if cfg_default.attachment_list is not None:
            self._column_cfg.attachment_list = 0
        return True

    def open_excel(self, xls_file: str) -> int:
        ret = self._excel.open_book(xls_file, False)
        if not ret:
            return False
        self._sheets_name = self._excel.sheets_name()
        return True

    def _load_sheet(self) -> bool:
        """

        :return: bool
        """
        # 读取EXCEL Sheet的信息
        (self._sheet_row_start, self._sheet_column_start, self._sheet_row_end, self._sheet_column_end,
         self._sheet_data) = self._excel.used_range_data()
        #
        self._row_cfg.caption = self._sheet_row_start
        self._row_cfg.send_start = self._sheet_row_start + 1
        self._row_cfg.send_end = self._sheet_row_end
        self._row_cfg.mail_count = self._row_cfg.send_end - self._row_cfg.send_start + 1
        #
        if self._column_cfg.sender != 0:
            self._column_cfg.sender += self._sheet_column_start
        if self._column_cfg.to != 0:
            self._column_cfg.to += self._sheet_column_start - 1
        if self._column_cfg.cc != 0:
            self._column_cfg.cc += self._sheet_column_start - 1
        if self._column_cfg.subject != 0:
            self._column_cfg.subject += self._sheet_column_start - 1
        if self._column_cfg.body != 0:
            self._column_cfg.body += self._sheet_column_start - 1
        if self.check_run_para:
            return False
        self._load_sheet_success = True
        return True

    def load_sheet_byname(self, sheet_name: str) -> bool:
        ret = self._excel.load_sheet_byname(sheet_name)
        if not ret:
            return False
        ret = self._load_sheet()
        if not ret:
            return False
        return True

    def load_sheet_byindex(self, index: int) -> bool:
        ret = self._excel.load_sheet_byindex(index)
        if not ret:
            return False
        ret = self._load_sheet()
        if not ret:
            return False
        return True

    @property
    def check_run_para(self) -> bool:
        """
        :return:
        """
        if self._row_cfg.send_start <= 0 or self._row_cfg.send_end <= 0:
            return False
        if self._column_cfg.sender > 0 or self._column_cfg.to:
            return False
        return True

    def send_one(self, line: int) -> bool:
        """
        发送一封邮件
        :param line: 发送的行号
        :rtype: bool
        """
        new_mail = copy.copy(self._default_mail)

        if 0 != self._column_cfg.sender:
            new_mail.sender = self._sheet_data[line - 1][self._column_cfg.sender - 1]
        if 0 != self._column_cfg.to:
            new_mail.to = self._sheet_data[line - 1][self._column_cfg.to - 1]
        if 0 != self._column_cfg.cc:
            new_mail.cc = self._sheet_data[line - 1][self._column_cfg.cc - 1]
        if 0 != self._column_cfg.subject:
            new_mail.subject = self._sheet_data[line - 1][self._column_cfg.subject - 1]
        if 0 != self._column_cfg.body:
            new_mail.body = self._sheet_data[line - 1][self._column_cfg.body - 1]

        if not self._column_cfg.attachment_list:
            for column_attachment in self._column_cfg.attachment_list:
                attach = self._sheet_data[line - 1][column_attachment - 1]
                new_mail.attachment_list.append(attach)

        self._outlook.create_sendmail()
        self._outlook.set_send_account(new_mail.sender)
        self._outlook.set_sendmail_all(new_mail)
        self._outlook.send_mail()
        return True

    def send_all(self) -> bool:
        i = 0
        while i < self._row_cfg.mail_count:
            send_result = self.send_one(i)
            if not send_result:
                return False
            i += 1
            time.sleep(0.1)
        return True

    @property
    def sheets_name(self) -> list:
        # 打开的EXCEL的sheets的名称列表
        return self._sheets_name

    def cur_sheet_info(self):
        # 当前的sheet的信息，起始和结束行列
        # load 后使用
        return self._sheet_row_start, self._sheet_column_start, self._sheet_row_end, self._sheet_column_end

    @property
    def mail_count(self) -> int:
        return self._row_cfg.mail_count

    def close(self):
        self._excel.quit()
        self._outlook.quit()


if __name__ == '__main__':
    pass
