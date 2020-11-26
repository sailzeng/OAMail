import win32com.client as win32

class DataProcessExcel(object):
    def __init__(self):
        self.excel_app_ = None
        self.sheets_ = None
        self.sheet_ = None

    def start(self):
        try:
            self.excel_app_ = win32.Dispatch('Excel.Application')
        except Exception as value:
            print("Exception occured, value = ", value)
        return

    def quit(self):
        '''
        退出EXCEL,
        :return:
        '''
        self.excel_app_.Quit()
        return

if __name__ == '__main__':
    print("Hello world!")