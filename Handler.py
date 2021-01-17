import imaplib
import getpass
import email
import csv
from imapclient import IMAPClient
import datetime
from itertools import chain

imaplib._MAXLINE = 1000000
class EmailHandler:

    def __init__(self):

        self.server = IMAPClient('imap.gmail.com', use_uid=True, ssl=True)


    def login(self, username, password):
        """

        :param username:
        :param password:
        :return:
        """
        try:
            self.server.login(username, password)
            print("Authentication Success")
        except IMAPClient.Error as e:
            print(e)

    def open_folder(self, folder):
        """

        :param folder:
        :return:
        """
        try:
            self.folder = self.server.select_folder(folder)
            print("%d messages in %s "%(self.folder[b'EXISTS'], folder))
        except IMAPClient.Error as e:
            print(e)
    def search(self, search_term):
        """
        search_term: List
                    a search term to retrieve specific email messages that satisfy term
        """
        try:
            if len(search_term) > 1:
                print("Fetching messages from %s.."%search_term[1])
            else:
                print("Fetching messages...")
            messages = self.server.search(search_term)
            return messages
        except IMAPClient.Error as e:
            print(e)

    def get_messages(self, messages):
        """

        :return:
        """
        print("Formatting messages..")
        self.dictionary = {}
        for uid, message_data in self.server.fetch(messages, ['ENVELOPE']).items():
            envelope = message_data[b'ENVELOPE']
            email_date = envelope.date
            address = envelope.from_[0]
            sender_mail = str(address.mailbox, 'utf-8') + "@" + str(address.host, 'utf-8')
            if address.name != None: #Some older emails wouldn't have a name property
                # sender_name = str(address.name, 'utf-8')
                sender_name = email.header.decode_header(str(address.name, 'utf-8'))[0][0]
                if isinstance(sender_name, bytes):
                    sender_name = str(sender_name, 'utf-8')
            else:
                sender_name = sender_mail

            if envelope.subject != None:
                subject = email.header.decode_header(str(envelope.subject, 'utf-8'))[0][0]
                if isinstance(subject, bytes):
                    subject = str(subject, 'utf-8')
            else:
                subject = 'None'
            if sender_mail not in self.dictionary:
                # key: sender email
                # values: [sender name, email count, oldest date from sender, newest date from sender]
                self.dictionary[sender_mail] = [sender_name, 1, email_date, email_date,subject]
            else:
                self.dictionary[sender_mail][1] += 1 # increase email count
                # self.dictionary[sender_mail][4].append(subject)
                if self.dictionary[sender_mail][2] > email_date: # adjusting oldest date
                    self.dictionary[sender_mail][2] = email_date
                if self.dictionary[sender_mail][3] < email_date: #adjusting newest date
                    self.dictionary[sender_mail][3] = email_date



    def create_csv(self, report_name):
        """

        :param report_name:
        :return:
        """
        print("creating report..")
        now = datetime.datetime.now().strftime("%d-%m-%Y")
        file_name = report_name + now + ".csv"
        with open(file_name, mode='w', newline="", encoding='utf-8-sig') as csvfile:
            writer = csv.writer(csvfile, delimiter=",", quotechar='"', quoting=csv.QUOTE_MINIMAL)
            writer.writerow(["Name", "Email", "Count", "Oldest Email", "Newest Email","Subjects"])
            for key, value in self.dictionary.items():
                writer.writerow([value[0],key, value[1], value[2], value[3], value[4]])

    def read_csv(self, file):
        """

        :param file:
        :return:
        """
        print("reading csv")
        UIDs = []
        with open(file, mode='r') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                if row['delete'] == "1":
                    address = row['Email']
                    UIDs.append((self.search(["FROM", address])))
        UIDs = list(chain.from_iterable(UIDs))
        return UIDs



handler = EmailHandler()
handler.login(username = input("Username: "), password=getpass.getpass(prompt='Password: '))
# handler.login(username, password)

# search_term1 = ['FROM', 'notify@twitter.com']
# search_term2 = ['SINCE', datetime.date(2019, 6, 1)]

# search_result1 = handler.search(search_term1)
# search_result2 = handler.search(search_term2)
# res = list(set(search_result1).intersection(search_result2))
# print(set(res1).intersection(res2))
# print(res)

# ========================================


handler.open_folder("inbox")
create_report = input("Create Report Y/N:  ")
if create_report.lower() == "y":
    search_term = ['ALL']
    search_result = handler.search(search_term)
    if len(search_result) != 0:
        handler.get_messages(search_result)
        handler.create_csv("email_reports ")
    else:
        print("No messages")
delete_file = 'delete_report.csv'
read_csv = input("Would you like to read csv specified for deletion?")
if read_csv.lower() == 'y':
    UIDs = handler.read_csv(delete_file)
    delete = input("%d emails selected for deletion, would you like to proceed?" %len(UIDs))
    if delete.lower() == 'y':
        pass


