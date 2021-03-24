import unittest
import batch_send as bps


class UTBatchProcessSend(unittest.TestCase):
    """

    """

    def test_batchsend(self):
        batch_send = bps.BatchProcessSend()
        ret = batch_send.start()
        self.assertEqual(ret, True)
        if not ret:
            return ret
        batch_send.config_column(0, 1, 2, 3, 4)
        batch_send.config_row(1, 2, 3)
        batch_send.config_default("insail@163.com")
        ret = batch_send.open_excel("D:\\SendMail.2.xlsx")
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