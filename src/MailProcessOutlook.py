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
        return

    # 清理所有的输出信息
    def clear_out_info(self):
        self.out_info_.clear()
        return

    # 独立初始化函数
    def init_outlook(self):
        try:
            # 使用MAPI协议连接Outlook
            self.outlook_ = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')
            self.outlook_.Visible = 0
            self.outlook_.DisplayAlerts = 0
            # 读取帐号
            self.accounts_ = self.outlook_.Session.Accounts
        except Exception as value:
            print("Exception occured, value = ", value)
        return

    def send_mail(self, send_account):
        # 0: olMailItem
        mail_item = self.outlook_.CreateItem(0)

        mail_item.Recipients.Add('someone@qq.com')
        mail_item.Subject = 'Mail Test'

        # 2: Html format
        mail_item.BodyFormat = 2
        mail_item.HTMLBody = '''
            <H2>Hello, This is a test mail.</H2>
            Hello Guys. 
            '''
        mail_item.Send()
        return

    """连接Outlook邮箱，读取收件箱内的邮件内容"""

    def read_outlook_mailbox(self, read_account, folder, read_num):

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
            ol_folder = 6

        read_box = None
        # 如果不指定账号，读取默认邮箱
        if read_account == "":
            read_box = self.outlook_.GetDefaultFolder(ol_folder)
        else:
            for account in self.accounts_:
                if read_account == account.DeliveryStore.DisplayName:
                    read_box = self.outlook_.Folders(account.DeliveryStore.DisplayName)
                    print("****Account Name**********************************", file=f)

        if not read_box:
            return -1

        # 获取收件箱下的所有邮件
        mails = read_box.Items
        mails.Sort('[ReceivedTime]', True)  # 邮件按时间排序
        # 读取收件箱内前3封邮件的所有信息（下标从1开始）
        for index in range(1, 4):
            print('正在读取第[{}]封邮件...'.format(index))
            mail = mails.Item(index)
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

            # 保存邮件中的附件，如果没有附件不会执行也不会产生异常
            attachment = mail.Attachments
            for each in attachment:
                save_attachment_path = os.getcwd()  # 保存附件到当前路径
                each.SaveAsFile(r'{}\{}'.format(save_attachment_path, each.FileName))
                print('附件（{}）保存完毕'.format(each.FileName))

        return


if __name__ == '__main__':
    send_mail()
