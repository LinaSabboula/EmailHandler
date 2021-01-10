import imaplib
import getpass
import email
import webbrowser
from imapclient import IMAPClient
from collections import defaultdict
from datetime import date

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
        messages = server.search('ON 8-JAN-2021')
        # messages = server.search('ALL')
        dictionary = {}
        for uid, message_data in server.fetch(messages, ['ENVELOPE']).items():
            envelope = message_data[b'ENVELOPE']
            address = envelope.from_[0]
            sender_name = str(address.name, 'utf-8')

            sender_mail = str(address.mailbox, 'utf-8') + "@" + str(address.host, 'utf-8')
            subject = email.header.decode_header(str(envelope.subject, 'utf-8'))[0][0]
            email_date = envelope.date
            if isinstance(subject, bytes):
                subject = str(subject, 'utf-8')
            if sender_mail not in dictionary:
                # key: sender email
                # values: [sender name, email count, oldest date from sender, newest date from sender]
                dictionary[sender_mail] = [sender_name, 1, email_date, email_date]
            else:
                dictionary[sender_mail][1] += 1 # increase email count
                if dictionary[sender_mail][2] > email_date: # adjusting oldest date
                    dictionary[sender_mail][2] = email_date
                if dictionary[sender_mail][3] < email_date: #adjusting newest date
                    dictionary[sender_mail][3] = email_date

    except IMAPClient.Error as e:
        print(e)

except IMAPClient.Error as e:
    print(e)
