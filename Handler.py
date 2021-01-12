import imaplib
import getpass
import email
import csv
from imapclient import IMAPClient
import datetime


imaplib._MAXLINE = 1000000
server = IMAPClient('imap.gmail.com', use_uid=True, ssl=True)
username = input("Username: ")

try:
    server.login(username, getpass.getpass(prompt="Password: "))
    # server.login(username, password)
    print("Authentication Success")
    folder = server.select_folder('INBOX')
    print("%d messages in INBOX "%folder[b'EXISTS'])
    try:
        # messages = server.search('ON 8-JAN-2021')
        messages = server.search('ALL')
        # messages = server.search(['FROM', 'notify@twitter.com'])
        dictionary = {}
        for uid, message_data in server.fetch(messages, ['ENVELOPE']).items():
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
            if sender_mail not in dictionary:
                # key: sender email
                # values: [sender name, email count, oldest date from sender, newest date from sender]
                dictionary[sender_mail] = [sender_name, 1, email_date, email_date,[subject]]
            else:
                dictionary[sender_mail][1] += 1 # increase email count
                dictionary[sender_mail][4].append(subject)
                if dictionary[sender_mail][2] > email_date: # adjusting oldest date
                    dictionary[sender_mail][2] = email_date
                if dictionary[sender_mail][3] < email_date: #adjusting newest date
                    dictionary[sender_mail][3] = email_date


        now = datetime.datetime.now().strftime("%d-%m-%Y")
        file_name = "email_reports " + now + ".csv"
        with open(file_name, mode='w', newline="", encoding='utf-8-sig') as csvfile:
            writer = csv.writer(csvfile, delimiter=",", quotechar='"', quoting=csv.QUOTE_MINIMAL)
            writer.writerow(["Name", "Email", "Count", "Oldest Email", "Newest Email","Subjects"])
            for key, value in dictionary.items():
                writer.writerow([value[0],key, value[1], value[2], value[3], value[4]])


    except IMAPClient.Error as e:
        print(e)

except IMAPClient.Error as e:
    print(e)
