import os
import win32com.client as win32
import enum


# 邮件处理，处理OUTLOOK
class MailProcessOutlook(object):
    @enum.unique
    class OutlookFolder(enum.Enum):
        FOLDER_OUTBOX = 1
        FOLDER_SENDMAIL = 2
        FOLDER_INBOX = 3
        FOLDER_DRAFTS = 4

    def __init__(self):
        self.outlook_ = None
        self.accounts_ = None
        self.current_account_ = None
        self.out_info_ = []
        self.read_folder_ = None
        self.read_mails_ = None

    # 清理所有的输出信息
    def clear_out_info(self):
        self.out_info_.clear()
        return

    # 独立初始化函数
    def start(self):
        try:

            # 使用MAPI协议连接Outlook
            self.outlook_ = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')
            # 读取帐号
            self.accounts_ = self.outlook_.Session.Accounts
            # 还有一种写法
            # self.outlook_ =  win32.gencache.EnsureDispatch('Excel.Application')
            # self.outlook_.Visible = 0
            self.outlook_.DisplayAlerts = 0

        except Exception as value:
            print("Exception occured, value = ", value)
        return

    def quit(self):
        self.outlook_.Application.Quit()
        return

    def send_mail(self, sender, receiver, subject, body):
        send_account = None
        for account in self.accounts_:
            if account.DisplayName == sender:
                send_account = account
                break
        # 0: olMailItem
        mail_item = self.outlook_.CreateItem(0)

        mail_item.Recipients.Add(receiver)
        mail_item.Subject = subject
        mail_item.SendUsingAccount = send_account
        # 2: Html format
        mail_item.BodyFormat = 2
        mail_item.HTMLBody = body
        mail_item.Send()
        return

    def send_template_mail(self, sender, receiver, subject, template_mail):
        send_account = None
        if sender:
            for account in self.accounts_:
                if account.DisplayName == sender:
                    send_account = account
                    break
        else:
            send_account = self.accounts_.Item(1)
        ol = None
        ol = win32.Dispatch('Outlook.Application')
        # 0: olMailItem
        new_mail = ol.CreateItem(0)

        new_mail.Recipients.Add(receiver)
        new_mail.Subject = subject
        new_mail.SendUsingAccount = send_account
        # 2: Html format
        new_mail.BodyFormat = 2
        new_mail.HTMLBody = template_mail.HTMLBody
        # new_mail.Attachments = template_mail.Attachments
        new_mail.Send()
        return

    @staticmethod
    def print_mail(mail):
        print('接收时间：{}'.format(str(mail.ReceivedTime)[:-6]))
        print('发件人：{}'.format(mail.SenderName))
        print('收件人：{}'.format(mail.To))
        print('抄送人：{}'.format(mail.CC))
        print('主题：{}'.format(mail.Subject))
        print('邮件正文内容：{}'.format(mail.Body))
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

    """连接Outlook邮箱，读取收件箱内的邮件内容"""

    def read_folder(self, read_account, folder):

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

        self.read_folder_ = None
        # 如果不指定账号，读取默认邮箱
        if not read_account:
            self.read_folder_ = self.outlook_.GetDefaultFolder(ol_folder)
        else:
            for account in self.accounts_:
                if read_account == account.DeliveryStore.DisplayName:
                    self.read_folder_ = self.outlook_.Folders(account.DeliveryStore.DisplayName)
                    self.out_info_.append("Use account{}".format(account.DeliveryStore.DisplayName))
        if not self.read_folder_:
            return -1

        # 获取收件箱下的所有邮件
        self.read_mails_ = self.read_folder_.Items
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

        # 读取收件箱内前3封邮件的所有信息（下标从1开始）
        for index in range(1, 4):
            print('正在读取第[{}]封邮件...'.format(index))
            mail = self.read_mails_.Item(index)
            # 保存邮件中的附件，如果没有附件不会执行也不会产生异常
            MailProcessOutlook.print_mail(mail)
        return 0

    def sort_mail(self, property_name, descending):
        self.read_mails_.Sort(property_name, descending)
        return

    # 读取目录的邮件
    def read_folder_mail(self, filter_subject, max_read_num):
        read_num = 0
        read_list = []
        mail = None
        for mail in self.read_mails_:
            # 标题过滤
            if filter_subject and -1 == mail.Subject.find(filter_subject):
                continue
            read_list.append(mail)
            read_num += 1
            # 数量过滤
            if max_read_num > 0 and read_num < max_read_num:
                break
        return read_num, read_list

    def print_folder_mail(self, filter_subject, max_read_num):
        read_num = 0
        read_list = []
        mail = None
        for mail in self.read_mails_:
            # 标题过滤
            if filter_subject and -1 == mail.Subject.find(filter_subject):
                continue
            MailProcessOutlook.print_mail(mail)
            read_num += 1
            # 数量过滤
            if max_read_num > 0 and read_num < max_read_num:
                break
        return


if __name__ == '__main__':
    outlook = MailProcessOutlook()
    outlook.start()
    outlook.read_folder(None, MailProcessOutlook.OutlookFolder.FOLDER_DRAFTS)
    outlook.print_folder_mail(None, 3)
    mail_num, mail_list = outlook.read_folder_mail("如何", 3)
    print("Mail number {}".format(mail_num))
    outlook.send_template_mail(None, "someone@qq.com" ,"TEST",mail_list[0])
    # outlook.quit()
