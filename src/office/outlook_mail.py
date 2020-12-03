import os
import win32com.client as win32
import enum


# 邮件处理，处理OUTLOOK
class OutlookMail(object):
    @enum.unique
    class OutlookFolder(enum.Enum):
        FOLDER_OUTBOX = 1
        FOLDER_SENDMAIL = 2
        FOLDER_INBOX = 3
        FOLDER_DRAFTS = 4

    def __init__(self):
        self._outlook_app = None
        self._ol_namespace = None
        self._accounts = None
        self._send_account = None
        self._out_info = []
        self._read_folder = None
        self._read_mails = None

    # 清理所有的输出信息
    def clear_out_info(self):
        self._out_info.clear()
        return

    # 独立初始化函数
    def start(self):
        try:
            # 使用MAPI连接Outlook
            self._outlook_app = win32.Dispatch('Outlook.Application')
            self._ol_namespace = self._outlook_app.GetNamespace('MAPI')
            # 读取帐号
            self._accounts = self._ol_namespace.Session.Accounts
            # 还有一种写法
            # self._outlook_app =  win32.gencache.EnsureDispatch('Excel.Application')
            # self._outlook_app.Visible = 0
            self._ol_namespace.DisplayAlerts = 0
        except Exception as value:
            print("Exception occured, value = ", value)
        return

    def quit(self):
        '''
        退出Outlook,
        :return:
        '''
        self._ol_namespace.Application.Quit()
        return

    def get_accounts(self):
        """
        返回OUTLOOK账户列表，相关数据查询Account object
        :return:
        """
        accounts = []
        for account in self._accounts:
            accounts.append(account)
        return accounts

    def set_send_account(self, sender):
        """
        设置默认的发送人账户，
        :param sender:
        :return:
        """
        if sender:
            for account in self._accounts:
                if account.DisplayName == sender:
                    self._send_account = account
                    break
        else:
            self._send_account = self._accounts.Item(1)
        return

    def send_mail(self, receiver, subject, body):
        """
        发送邮件
        :param receiver: 邮件接受者
        :param subject: 邮件标题
        :param body: 邮件内容
        :return: 无
        """
        # 0: olMailItem
        mail_item = self._outlook_app.CreateItem(0)

        mail_item.Recipients.Add(receiver)
        mail_item.Subject = subject
        mail_item.SendUsingAccount = self._send_account
        # 2: Html format
        mail_item.BodyFormat = 2
        mail_item.HTMLBody = body
        mail_item.Send()
        return

    def send_mail_from_copy(self, receiver, subject, copy_mail):
        """
        根据模版邮件发送

        """
        # 0: olMailItem
        new_mail = self._outlook_app.CreateItem(0)
        # 2: Html format
        new_mail.BodyFormat = 2
        new_mail = copy_mail.Copy()
        new_mail.Recipients.Add(receiver)
        new_mail.Subject = subject
        new_mail.SendUsingAccount = self._send_account
        new_mail.Send()
        return

    def read_account_folder(self, read_account, folder):
        """
        读取OUTLOOK的某个帐号的folder，（收件箱等）
        :param read_account: 读取帐号
        :param folder: OutlookFolder.FOLDER_OUTBOX枚举值
        :return: True，成功，False 失败
        """
        # OLE 的很多枚举值
        # olFolderDeletedItems 3 已发送
        # olFolderOutbox 4 发件箱
        # olFolderSentMail 5 已经发送邮件
        # olFolderInbox 6 收件箱
        # olFolderDrafts 16 草稿箱
        ol_folder = 0
        if folder == self.OutlookFolder.FOLDER_OUTBOX:
            ol_folder = 3
        elif folder == self.OutlookFolder.FOLDER_SENDMAIL:
            ol_folder = 4
        elif folder == self.OutlookFolder.FOLDER_INBOX:
            ol_folder = 6
        elif folder == self.OutlookFolder.FOLDER_DRAFTS:
            ol_folder = 16
        else:
            assert False

        self._read_folder = None
        # 如果不指定账号，读取默认邮箱
        if not read_account:
            self._read_folder = self._ol_namespace.GetDefaultFolder(ol_folder)
        else:
            for account in self._accounts:
                if read_account == account.DeliveryStore.DisplayName:
                    self._read_folder = self._ol_namespace.Folders(account.DeliveryStore.DisplayName)
                    self._out_info.append("Use account{}".format(account.DeliveryStore.DisplayName))
        if not self._read_folder:
            return -1
        # 获取收件箱下的所有邮件
        self._read_mails = self._read_folder.Items
        # 排序，不同的邮箱用不同的排序方式
        if folder == self.OutlookFolder.FOLDER_OUTBOX:
            self.sort_mail('[SentOn]', True)
        elif folder == self.OutlookFolder.FOLDER_SENDMAIL:
            self.sort_mail('[SentOn]', True)
        elif folder == self.OutlookFolder.FOLDER_INBOX:
            self.sort_mail('[ReceivedTime]', True)
        elif folder == self.OutlookFolder.FOLDER_DRAFTS:
            self.sort_mail('[LastModificationTime]', True)
        else:
            assert False
        return 0

    def sort_mail(self, property_name, descending):
        '''

        :param property_name:
        :param descending:
        :return:
        '''
        self._read_mails.Sort(property_name, descending)
        return

    def read_folder_mail(self, filter_subject, max_read_num):
        """
        读取folder的邮件列表，前面应该先调用read_folder
        :param filter_subject: 过滤标题的文本
        :param max_read_num: 最大的读取邮件数量
        :return: 读取的邮件列表
        """
        read_num = 0
        read_list = []
        mail = None
        for mail in self._read_mails:
            # 标题过滤
            if filter_subject and -1 == mail.Subject.find(filter_subject):
                continue
            read_list.append(mail)
            read_num += 1
            # 数量过滤
            if max_read_num > 0 and read_num < max_read_num:
                break
        return read_list

    def print_folder_mail(self, filter_subject, max_read_num):
        read_num = 0
        read_list = []
        mail = None
        for mail in self._read_mails:
            # 标题过滤
            if filter_subject and -1 == mail.Subject.find(filter_subject):
                continue
            OutlookMail.print_mail(mail)
            read_num += 1
            # 数量过滤
            if max_read_num > 0 and read_num < max_read_num:
                break
        return

    def print_accounts(self):
        accounts = self.get_accounts()
        num = 0
        for account in accounts:
            print('Account No：{}'.format(num))
            OutlookMail.print_account(account)
            num += 1
        return

    @staticmethod
    def print_mail(mail):
        print('ReceivedTime：{}'.format(str(mail.ReceivedTime)[:-6]))
        print('SenderName：{}'.format(mail.SenderName))
        print('To：{}'.format(mail.To))
        # 有的邮件没有CC
        print('CC：{}'.format(mail.CC))
        print('Subject：{}'.format(mail.Subject))
        print('Body：{}'.format(mail.Body))
        print('邮件附件数量：{}'.format(mail.Attachments.Count))
        print('邮件MessageID：{}'.format(mail.EntryID))
        print('会话主题：{}'.format(mail.ConversationTopic))
        print('会话ID：{}'.format(mail.ConversationID))
        print('会话记录相对位置：{}'.format(mail.ConversationIndex))
        attachment = mail.Attachments
        for each in attachment:
            # save_attachment_path = os.getcwd()  # 保存附件到当前路径
            # each.SaveAsFile(r'{}\{}'.format(save_attachment_path, each.FileName))
            print('附件（{}）保存完毕'.format(each.FileName))
        return

    @staticmethod
    def print_account(account):
        print('AccountType：{}'.format(account.AccountType))
        print('CurrentUser ：{}'.format(account.CurrentUser))
        print('DisplayName ：{}'.format(account.DisplayName))
        print('UserName：{}'.format(account.UserName))
        return


if __name__ == '__main__':
    outlook = OutlookMail()
    outlook.start()
    outlook.read_account_folder(None, OutlookMail.OutlookFolder.FOLDER_DRAFTS)
    outlook.print_folder_mail(None, 3)
    mail_list = outlook.read_folder_mail(None, 3)
    print("Mail number {}".format(len(mail_list)))
    outlook.print_accounts()
    # outlook.set_send_account(None)
    # outlook.send_mail_from_copy("halozhao@tencent.com", "TEST", mail_list[0])
    # outlook.quit()
