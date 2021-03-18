import win32com.client as win32
import os


class ExcelDataBase(object):
    def __init__(self):
        self._excel_app = None
        self._work_book = None
        self._work_sheets = None
        self._work_sheet = None
        self._sheet = None
        self._xls_file = None
        self._is_open = False
        self._is_new = False

    def start(self) -> bool:
        try:
            self._excel_app = win32.Dispatch('Excel.Application')
        except Exception as value:
            print("Exception occured, value = ", value)
            return False
        return True

    def quit(self):
        """
        退出EXCEL,
        :return:
        """
        self._excel_app.Quit()
        return

    def open_book(self, file_name, not_exist_new) -> bool:
        """        """
        # 如果指向的文件不存在，则需要新建一个
        if os.path.exists(file_name) and os.path.isfile(file_name):
            if not not_exist_new:
                return False

        # 得到绝对路径，因为ActiveX只支持绝对路径，包括Open，包括 Save as,
        # 不光必须用绝对路径，还需要实用原生的路径分割符号'\'

        self._work_book = self._excel_app.Workbooks.Open(file_name)
        # self.work_book_ = self.excel_app_.ActiveWorkBook
        self._xls_file = os.path.abspath(file_name)
        self._is_open = True
        self._work_sheets = self._work_book.Worksheets
        return True

    def new_book(self):
        """        """
        # 新建一个xls，添加一个新的工作薄
        self._excel_app.WorkBooks.Add()
        self._work_book = self._excel_app.ActiveWorkBook
        self._is_open = True
        self._is_new = True
        return True

    def close(self):
        """关闭打开的EXCEL,"""
        self._is_open = False
        self._xls_file = None
        self._work_book.Close(True)
        self._is_open = False
        self._is_new = False
        return

    def sheets_count(self):
        count = self._work_sheets.Count
        return count

    def sheets_name(self):
        name_list = []
        count = self._work_sheets.Count
        i = 0
        while i < count:
            name_list.append(self._work_book.Worksheets(i + 1).Name)
        return name_list

    def load_sheet_byindex(self, sheet_index: int):
        self._work_sheet = self._work_book.Worksheets(sheet_index)
        if not self._work_book:
            return False
        return True

    def load_sheet_byname(self, sheet_name: str):
        """

        :param sheet_name: 
        :return: 
        """
        self._work_sheet = self._work_book.Worksheets(sheet_name)
        if not self._work_book:
            return False
        return True

    @staticmethod
    def _range_coord(read_range):
        """"""
        row_count = read_range.Rows.Count
        column_count = read_range.Columns.Count
        # 因为excel可以从任意行列填数据而不一定是从1, 1 开始，因此要获取首行列下标
        # 第一行,列的起始位置
        row_start = read_range.Row
        column_start = read_range.Column
        return row_start, column_start, row_count, column_count

    @staticmethod
    def _range_data(read_range):
        """"""
        if not read_range:
            return 0, 0, 0, 0, []
        else:
            row_start, column_start, row_count, column_count = \
                ExcelDataBase._range_coord(read_range)
            ret_list = []
            i = 0
            j = 0
            while i < row_count:
                line = []
                while j < column_count:
                    line.append(read_range.Cell(i, j))
                    j += 1
                ret_list.append(line)
                i += 1
        return row_start, column_start, row_count, column_count, ret_list

    def used_range_coord(self):
        """
        取得UsedRange的各种坐标，包括起始行号，列号，以及占用的行总数，列总数
        注意UsedRange并不一定是从0，0开始的
        :return:UsedRange的行起始，列起始，行总数，列总数
        """
        used_range = self._work_sheet.UsedRange
        if not used_range:
            return 0, 0, 0, 0
        else:
            return self._range_coord(used_range)

    def used_range_data(self) -> object:
        used_rg = self._work_sheet.UsedRange
        return self._range_data(used_rg)

    def sheet_cell(self, row, column):
        # 如果预加载了数据,
        return self._work_sheet.Cells(row, column).Value

    @staticmethod
    def column_name(column_num):
        """"""
        assert column_num > 0
        n = column_num
        lst = []
        while True:
            if n > 0:
                # EXCEL 奇特的规则导致的这个地方，没有0，和一般的转码不太一样
                n -= 1
            m = n % 26
            n //= 26
            lst.append(chr(m + ord('A')))
            if n <= 0:
                break
        lst.reverse()
        return "".join(lst)

    def range(self, cell1: str, cell2: str = None):
        return self._work_sheet.Range(cell1, cell2)

    def range2(self, cell1_row: int, cell1_column: int, cell2_row: int, cell2_column: int):
        cell1 = str(cell1_row) + ExcelDataBase.column_name(cell1_column)
        cell2 = str(cell2_row) + ExcelDataBase.column_name(cell2_column)
        return self._work_sheet.Range(cell1, cell2)

    def range_coord(self, cell1: str, cell2: str = None):
        get_range = self.range(cell1, cell2)
        return self._range_coord(get_range)

    def range_data(self, cell1: str, cell2: str = None):
        get_range = self.range(cell1, cell2)
        return self._range_data(get_range)


if __name__ == '__main__':
    print("Hello world!{}".format(__file__))
    print("column_name {} {}".format(1, ExcelDataBase.column_name(1)))
    print("column_name {} {}".format(26, ExcelDataBase.column_name(26)))
    print("column_name {} {}".format(27, ExcelDataBase.column_name(27)))
    print("column_name {} {}".format(52, ExcelDataBase.column_name(52)))
    print("column_name {} {}".format(200, ExcelDataBase.column_name(200)))
    print("column_name {} {}".format(888, ExcelDataBase.column_name(888)))
