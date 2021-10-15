import unittest
import batch_send as bps
import office.outlook_mail as outlook


class UTBatchProcessSend(unittest.TestCase):
    """

    """

    def test_batch_send(self):
        batch_send = bps.BatchProcessSend()
        ret = batch_send.start()
        self.assertEqual(ret, True)
        if not ret:
            return ret

        ret = batch_send.open_excel("D:\\SendMail.2.xlsx")
        self._default_mail = outlook.SendMailInfo()
        self._row_config = bps.RowNumConfig()
        self._column_config = bps.ColumnNumConfig()
        self._default_mail.sender = "insail@163.com"
        self._column_config.sender = 0
        self._column_config.to = 1
        self._column_config.cc = 2
        self._column_config.subject = 3
        self._column_config.body = 4
        self._row_config.caption = 1
        self._row_config.send_start = 2
        self._row_config.send_end = 3

        batch_send.config(self._column_config,
                          self._row_config,
                          self._default_mail)
        self.assertEqual(ret, True)
        if not ret:
            return ret
        ret = batch_send.load_sheet_byindex(1)
        self.assertEqual(ret, True)
        if not ret:
            return ret
        ret = batch_send.send_all()
        self.assertEqual(ret, True)
        if not ret:
            return ret
        return True


if __name__ == '__main__':
    unittest.main()
