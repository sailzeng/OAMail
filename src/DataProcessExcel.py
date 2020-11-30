import win32com.client as win32
import os


class DataProcessExcel(object):
    def __init__(self):
        self.excel_app_ = None
        self.work_book_ = None
        self.work_sheets_ = None
        self.work_sheet_ = None
        self.sheet_ = None
        self.xls_file_ = None
        self.is_open_ = False
        self.is_new_ = False

    def start(self):
        try:
            self.excel_app_ = win32.Dispatch('Excel.Application')
        except Exception as value:
            print("Exception occured, value = ", value)
        return

    def quit(self):
        """
        退出EXCEL,
        :return:
        """
        self.excel_app_.Quit()
        return

    def open_book(self, file_name, not_exist_new):
        """        """
        # 如果指向的文件不存在，则需要新建一个
        if os.path.exists(file_name) and os.path.isfile(file_name):
            if not not_exist_new:
                return False

        # 得到绝对路径，因为ActiveX只支持绝对路径，包括Open，包括Saveas,
        # 不光必须用绝对路径，还需要实用原生的路径分割符号'\'

        self.work_book_ = self.excel_app_.Workbooks.Open(file_name)
        # self.work_book_ = self.excel_app_.ActiveWorkBook
        self.xls_file_ = os.path.abspath(file_name)
        self.is_open_ = True
        self.work_sheets_ = self.work_book_.Worksheets
        return True

    def new_book(self):
        """        """
        # 新建一个xls，添加一个新的工作薄
        self.excel_app_.WorkBooks.Add()
        self.work_book_ = self.excel_app_.ActiveWorkBook
        self.is_open_ = True
        self.is_new_ = True
        return True

    def close(self):
        """关闭打开的EXCEL,"""
        self.is_open_ = False
        self.xls_file_ = None
        self.work_book_.Close(True)
        self.is_open_ = False
        self.is_new_ = False
        return

    def sheet_count(self):
        count = self.work_sheets_.Count
        return count

    def load_sheet(self, sheet_index: int, pre_read_data: bool):
        self.work_sheet_ = self.work_book_.Worksheets(sheet_index)
        if not self.work_book_:
            return False
        return True

    def load_sheet(self, sheet_name: str, pre_read_data: bool):
        self.work_sheet_ = self.active_book_.Worksheets(sheet_name)
        if not self.work_book_:
            return False
        return True

    def used_range_coord(self):
        """
        取得UsedRange的各种坐标，包括起始行号，列号，以及占用的行总数，列总数
        注意UsedRange并不一定是从0，0开始的
        :return:UsedRange的行起始，列起始，行总数，列总数
        """
        used_range = self.work_sheet_.UsedRange
        if not used_range:
            return 0, 0, 0, 0
        else:
            row_count = used_range.Rows.Count
            column_count = used_range.Columns.Count
            # 因为excel可以从任意行列填数据而不一定是从1, 1 开始，因此要获取首行列下标
            # 第一行,列的起始位置
            row_start = used_range.Row
            column_start = used_range.Column
            return row_start, column_start, row_count, column_count

    def used_range(self):
        used_range = self.work_sheet_.UsedRange
        if not used_range:
            return 0, 0, 0, 0, []
        else:
            row_start, column_start, row_count, column_count = self.used_range_coord()
            ret_list = []
            i = 0
            j = 0
            while i < row_count:
                line = []
                while j < column_count:
                    line.append(used_range.Cell(i, j))
                    j += 1
                ret_list.append(line)
                i += 1
        return row_start, column_start, row_count, column_count, ret_list

    def sheet_cell(self, row, column):
        # 如果预加载了数据,
        return self.work_sheet_.Cells(row, column).Value

    @staticmethod
    def column_name(column_num):
        """"""
        assert column_num > 0
        n = column_num
        lst = []
        while True:
            if n > 0:
                # EXCEL 奇特的规则导致的这个地方，没有0
                n -= 1
            m = n % 26
            n //= 26
            lst.append(chr(m + ord('A')))
            if n <= 0:
                break
        lst.reverse()
        return "".join(lst)

    def get_range(self, cell1: str, cell2: str):
        return self.work_sheet_.Range(cell1, cell2)

    def range(self, cell1_row: int, cell1_column: int, cell2_row: int, cell2_column: int):
        cell1 = str(cell1_row) + DataProcessExcel.column_name(cell1_column)
        cell2 = str(cell2_row) + DataProcessExcel.column_name(cell2_column)
        return self.work_sheet_.Range(cell1, cell2)

    def range(self, cell1: str, cell2: str):
        get_range = self.range(cell1, cell2)
        row_count = get_range.Rows.Count
        column_count = get_range.Columns.Count
        row_start = get_range.Row
        column_start = get_range.Column
        return row_start, column_start, row_count, column_count

if __name__ == '__main__':
    print("Hello world!{}".format(__file__))
    print("column_name {} {}".format(1, DataProcessExcel.column_name(1)))
    print("column_name {} {}".format(26, DataProcessExcel.column_name(26)))
    print("column_name {} {}".format(27, DataProcessExcel.column_name(27)))
    print("column_name {} {}".format(52, DataProcessExcel.column_name(52)))
    print("column_name {} {}".format(200, DataProcessExcel.column_name(200)))
    print("column_name {} {}".format(888, DataProcessExcel.column_name(888)))