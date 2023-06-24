# Importing libraries
import imaplib
import email
import csv
import os

import yaml  # To load saved login credentials from a yaml file

with open("credentials.yml") as f:
    content = f.read()

# from credentials.yml import user name and password
my_credentials = yaml.load(content, Loader=yaml.FullLoader)

# Load the user name and password from the yaml file
user, password = my_credentials["user"], my_credentials["password"]

# URL for IMAP connection
imap_url = 'imap.gmail.com'

# Connection with GMAIL using SSL
my_mail = imaplib.IMAP4_SSL(imap_url)

# Log in using your credentials
my_mail.login(user, password)

# Select the Inbox to fetch messages
my_mail.select('Inbox')

# Define the number of emails to fetch (top 150 in this case)
num_emails_to_fetch = 2000

# Fetch the email IDs of the top num_emails_to_fetch emails
_, data = my_mail.search(None, 'ALL')
mail_id_list = data[0].split()[-num_emails_to_fetch:]

msgs = []  # Empty list to capture all messages

# Iterate through the email IDs and fetch the messages
for num in mail_id_list:
    typ, data = my_mail.fetch(num, '(RFC822)')
    msgs.append(data)
    # Process the messages and extract the desired information
    email_data = []  # List to store the extracted email data

    for msg in msgs[::-1]:
        for response_part in msg:
            if isinstance(response_part, tuple):
                email_message = email.message_from_bytes(response_part[1])
                title = email_message['subject']
                sender = email_message['from']
                sender_email = email.utils.parseaddr(email_message['from'])[1]
                date = email_message['date']

                email_data.append([title, sender, sender_email, date])

    #Define the CSV file path
    #csv_file_path = 'email_data.csv'
    if os.path.exists("email_data_2k.csv"):
        os.remove("email_data_2k.csv")

    csv_file_path='email_data_2k.csv'


    # Write the email data to the CSV file
    with open(csv_file_path, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['Title', 'Sender', 'Sender Email', 'Date'])  # Write the header row
        writer.writerows(email_data)  # Write the email data rows
