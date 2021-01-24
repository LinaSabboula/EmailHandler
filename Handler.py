import imaplib
import getpass
import email
import csv
from imapclient import IMAPClient
import datetime
from itertools import chain
from awesome_progress_bar import ProgressBar

imaplib._MAXLINE = 1000000
class EmailHandler:
    """
    Class that will summarize emails in a csv report and enables you to delete certain marked emails

    Attributes:
        server(IMAPClient): IMAPClient instance
        report_dictionary(dict): store relevant email information for reporting
    """
    def __init__(self):
        self.server = IMAPClient('imap.gmail.com', use_uid=True, ssl=True)
        self.report_dictionary = {}

    def login(self, username, password):
        """
        attempts to login to email provided the email address and password

        Args:
            username (str): email address
            password (str): email password

        """
        try:
            self.server.login(username, password)
            print("Authentication Success")
        except IMAPClient.Error as e:
            print(e)

    def open_folder(self, folder):
        """
        Opens certain selected folder from your email associated with server

        Args:
            folder (str): name of the folder

        Returns:
            folder[b'EXISTS'] (int): The number of emails in selected folder

        """
        try:
            folder = self.server.select_folder(folder)
            return folder[b'EXISTS']
        except IMAPClient.Error as e:
            print(e)
    def search(self, search_term):
        """
        Fetches emails that satisfy search string

        Args:
            search_term: (list((str)): a search term to retrieve specific email messages that satisfy term

        Returns:
            messages (list(int)): list of email unique identifiers (UID) that satisfy search term

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
        fetches the email messages and keeps track and formats the sender name, email address, count, oldest email,
        newest email, one subject to add in report_dictionary to use for report later

        Args:
            messages(List(int)): list of email unique identifiers (UID) that satisfy search term

        """
        print("Formatting messages..")
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
            if sender_mail not in self.report_dictionary:
                # key: sender email
                # values: [sender name, email count, oldest date from sender, newest date from sender, subject of email]
                # each email account will only have one subject value, just to quickly see what kind of emails are associated with address
                self.report_dictionary[sender_mail] = [sender_name, 1, email_date, email_date, subject]
            else:
                self.report_dictionary[sender_mail][1] += 1 # increase email count
                if self.report_dictionary[sender_mail][2] > email_date: # adjusting oldest date
                    self.report_dictionary[sender_mail][2] = email_date
                if self.report_dictionary[sender_mail][3] < email_date: #adjusting newest date
                    self.report_dictionary[sender_mail][3] = email_date


    def create_csv(self, report_name):
        """
        Creates a csv report called 'report_name' followed by today's date,
        with the data generated from the report_dictionary values from the get_messages function

        Args:
            report_name(str): the name for the csv report

        """
        print("creating report..")
        now = datetime.datetime.now().strftime("%d-%m-%Y")
        file_name = report_name + now + ".csv"
        with open(file_name, mode='w', newline="", encoding='utf-8-sig') as csvfile:
            writer = csv.writer(csvfile, delimiter=",", quotechar='"', quoting=csv.QUOTE_MINIMAL)
            writer.writerow(["Name", "Email", "Count", "Oldest Email", "Newest Email","Subjects"])
            for key, value in self.report_dictionary.items():
                writer.writerow([value[0],key, value[1], value[2], value[3], value[4]])

    def read_csv(self, file):
        """
        Reads a csv file and for each row where "Delete" column is marked as "1" and
        gets the UIDs for all emails sent from this email address

        Args:
            file(str): Name of the file to read data marked for deletion

        Returns:
             UIDs(list(int)): list of email unique identifiers (UID) that have been marked for deletion

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

    def delete(self, UIDs):
        """
        Move specified emails to Trash box

        Args:
            UIDs(list(int)): list of email unique identifiers (UID) that have been marked for deletion

        """
        total = len(UIDs)
        bar = ProgressBar(total, use_eta=True, bar_length=70, spinner_type='db')
        for message in UIDs:
            try:
                self.server.add_gmail_labels(message,'\Trash', silent=True)
                bar.iter(' Moved to Trash')
            except:
                bar.stop()


handler = EmailHandler()
handler.login(username = input("Username: "), password=getpass.getpass(prompt='Password: '))
# handler.login(username, password)

inbox_count = handler.open_folder("inbox")
print("%d messages in inbox "%inbox_count)
create_report = input("Create Report Y/N:  ")
if create_report.lower() == "y":
    search_term = ['ALL']
    search_result = handler.search(search_term)
    if len(search_result) != 0:
        handler.get_messages(search_result)
        handler.create_csv("email_reports - new")
    else:
        print("No messages")
delete_file = 'delete_report.csv'
read_csv = input("Would you like to read csv specified for deletion? Y/N: ")
if read_csv.lower() == 'y':
    UIDs = handler.read_csv(delete_file)
    delete = input("%d out of %d emails selected for deletion, would you like to proceed? Y/N: " %(len(UIDs), inbox_count))
    if delete.lower() == 'y':
        handler.delete(UIDs)

trash_count = handler.open_folder("[Gmail]/Trash")
print("%d emails are currently in trash"%trash_count)



