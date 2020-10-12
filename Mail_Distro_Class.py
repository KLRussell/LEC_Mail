from Mail_Distro_Settings import sql_engine, email_engine, attach_dir, tool
from KGlobal import ExchangeToMsg
from datetime import timedelta, datetime
from exchangelib import FileAttachment, ItemAttachment, EWSDateTime, Message, EWSTimeZone
from os.path import splitext, join, exists
from os import makedirs, remove, walk
from hashlib import md5
from bs4 import BeautifulSoup
from pandas import DataFrame
from traceback import format_exc


class MailDistro(object):
    __table = None
    __distro = None
    __to_emails = None
    __cc_emails = None
    __dirs = list()
    __files = list()
    __items = list()

    def process(self, distro, table):
        self.__distro = distro
        self.__table = table
        self.__dirs = list()
        self.__files = list()
        self.__items = list()

        try:
            tool.write_to_log("Starting process of %s" % self.__distro)

            self.__cleanup(write_to_log=False)
            max_date = self.__find_max_date()
            self.__grab_emails(max_date)
            self.__proc_email(self.__to_emails, 'To')
            self.__proc_email(self.__cc_emails, 'Cc')
            self.__upload_results()
        except Exception as e:
            tool.write_to_log(format_exc(), 'error')
            tool.write_to_log('Class process() function ran into an error code %s. Cleaning up' % type(e).__name__)
        finally:
            self.__cleanup()

    def __find_max_date(self):
        query = 'SELECT MAX(Date_Received) Date_Received FROM %s' % self.__table
        results = sql_engine.sql_execute(query_str=query)

        if results and results.results:
            return results.results[0]['Date_Received'].iloc[0]

    def __grab_emails(self, max_date):
        if max_date:
            tool.write_to_log('{0} => Retreiving e-mails from {1} till now'.format(self.__distro, max_date))
            date_received = max_date + timedelta(seconds=1)
            email_datetime = EWSDateTime(date_received.year, date_received.month,
                                         date_received.day, date_received.hour,
                                         date_received.minute,
                                         date_received.second, tzinfo=EWSTimeZone.timezone('UTC'))

            self.__to_emails = email_engine.inbox.children.filter(
                datetime_received__gt=email_datetime,
                display_to__contains=self.__distro
            )
            self.__cc_emails = email_engine.inbox.children.filter(
                datetime_received__gt=email_datetime,
                display_cc__contains=self.__distro
            )
        else:
            '''
            # This is to fill table when table is truncated else keep this section commented
            self.to_emails = self.account.inbox.children.filter(
                display_to__contains=lec_mail
            )
            self.cc_emails = self.account.inbox.children.filter(
                display_cc__contains=lec_mail
            )
            '''
            # Comment these two statements out if you uncomment the block above
            self.__to_emails = None
            self.__cc_emails = None

    def __handle_attach(self, message, attachments, received_date):
        file_ext = dict()

        for attachment in attachments:
            if not attachment.is_inline and isinstance(attachment, FileAttachment):
                file = splitext(attachment.name)

                if len(file) == 2:
                    ext = file[1]
                else:
                    ext = file[0]
            elif isinstance(attachment, ItemAttachment) and isinstance(attachment.item, Message):
                ext = '.msg'
            else:
                ext = None

            if ext:
                if ext in file_ext.keys():
                    file_ext[ext] += 1
                else:
                    file_ext[ext] = 1

        if len(file_ext) > 0:
            file_hash = md5(message.id.encode()).hexdigest()
            if received_date:
                local_dir = join(attach_dir, self.__distro, received_date.strftime("%Y%m%d"))
            else:
                local_dir = join(attach_dir, self.__distro, datetime.now().strftime("%Y%m%d"))

            file_name = '{0}.msg'.format(file_hash)
            local_path = join(local_dir, file_name)

            if local_dir not in self.__dirs:
                self.__dirs.append(local_dir)

            self.__files.append(local_path)

            if not exists(local_dir):
                makedirs(local_dir)

            msg_obj = ExchangeToMsg(message)
            msg_obj.save(local_path)

            return [str(file_ext), local_path]
        else:
            return [None, None]

    def __proc_email(self, email, email_type):
        if email:
            tool.write_to_log("{0} => Processing {1} '{2}' Emails".format(self.__distro, len(list(email)), email_type))

            for item in email:
                to = None
                cc = None
                to_emails = []
                cc_emails = []

                if item.to_recipients:
                    for to in item.to_recipients:
                        to_emails.append(to.email_address)

                    to = '; '.join(to_emails)

                if item.cc_recipients:
                    for cc in item.cc_recipients:
                        cc_emails.append(cc.email_address)

                    cc = '; '.join(cc_emails)

                if item.datetime_received:
                    date_received = item.datetime_received
                    date_received = date_received
                else:
                    date_received = None

                attach1 = None
                attach2 = None

                if item.attachments:
                    attachs = self.__handle_attach(item, item.attachments, date_received)

                    if attachs and len(attachs) == 2:
                        attach1 = attachs[0]
                        attach2 = attachs[1]

                if date_received:
                    date_received = date_received.strftime('%Y-%m-%d %H:%M:%S %p')

                if item.subject:
                    subject = item.subject
                else:
                    subject = None

                if item.text_body:
                    body = item.text_body
                elif item.body:
                    soup = BeautifulSoup(item.body, 'html.parser')
                    body = '\n'.join([span.text.strip() for span in soup.find_all('body')]).strip()

                    if not body:
                        body = item.body
                else:
                    body = None

                if item.sender and hasattr(item.sender, 'name'):
                    sender_name = item.sender.name
                else:
                    sender_name = None

                if item.sender and hasattr(item.sender, 'email_address'):
                    sender_email = item.sender.email_address
                else:
                    sender_email = None

                if item.importance:
                    importance = item.importance
                else:
                    importance = None

                if sender_name or sender_email:
                    self.__items.append([sender_name, sender_email, to, cc, importance, subject, body, attach1,
                                         attach2, date_received])
        else:
            tool.write_to_log("{0} => Processing 0 '{1}' Emails".format(self.__distro, email_type))

    def __upload_results(self):
        if self.__items:
            tbl = self.__table.split('.')
            df = DataFrame(self.__items, columns=['Sender_Name', 'Sender_Email', 'Email_To', 'Email_Cc',
                                                  'Importance', 'Subject', 'Body', 'Attachments', 'Attach_Path',
                                                  'Date_Received'])
            tool.write_to_log("{0} => Uploading {1} into ODS".format(self.__distro, len(df)))
            sql_engine.sql_upload(dataframe=df, table_name=tbl[1], table_schema=tbl[0], if_exists='append', index=None)

    def __cleanup(self, write_to_log=True):
        if write_to_log:
            tool.write_to_log("{0} => Cleaning up attachment directory".format(self.__distro))

        query = 'SELECT DISTINCT Attach_Path FROM %s WHERE Attach_Path IS NOT NULL' % self.__table

        results = sql_engine.sql_execute(query_str=query)

        if results and results.results:
            paths = results.results[0]

            if not paths.empty:
                for subdir, dirs, files in walk(join(attach_dir, self.__distro)):
                    for file in files:
                        file_found = False
                        filepath = join(subdir, file)

                        if filepath.endswith(".msg"):
                            for path in paths['Attach_Path'].tolist():
                                if path.lower() == str(filepath).lower():
                                    file_found = True
                                    break

                        if not file_found:
                            try:
                                remove(filepath)
                            finally:
                                pass
