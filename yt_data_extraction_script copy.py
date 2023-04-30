import imaplib
import email
import re
import pandas as pd

# URL for IMAP connection
imap_url = 'imap.gmail.com'

my_mail = imaplib.IMAP4_SSL(imap_url)

# Log in using your credentials
my_mail.login('email','application password')

# Select the Inbox to fetch messages
my_mail.select('Inbox')

#change the value to the email that you want to extract data from
key = 'FROM'
value = 'the email address that you want to extart the data from'
_, data = my_mail.search(None, key, value)  # Search for emails with specific key and value

mail_id_list = data[0].split()  # IDs of all emails that we want to fetch

msgs = []  # empty list to capture all messages
# Iterate through messages and extract data into the msgs list
for num in mail_id_list:
    typ, data = my_mail.fetch(num, '(RFC822)')  # RFC822 returns whole message (BODY fetches just body)
    msgs.append(data)

# Let us extract the right text and print on the screen

# Define a regular expression to match the required information
pattern = r"([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})|(\+972\d{8,9})"

# Define output file path
output_file_path = r'path\to\the\file.txt'

# Open the output file in write mode
with open(output_file_path, 'w') as f:
    # Loop over the messages in reverse order
    for msg in msgs[::-1]:
        for response_part in msg:
            if type(response_part) is tuple:
                my_msg = email.message_from_bytes((response_part[1]))
                for part in my_msg.walk():
                    # print(part.get_content_type())
                    if part.get_content_type() == 'text/plain':
                        # Find matches for the regular expression in the message body
                        matches = re.findall(pattern, part.get_payload())
                        if matches:
                            for match in matches:
                                # Check if the match is an email address or an Israeli phone number
                                email_match, phone_match = match
                                if email_match:
                                    f.write(f'Email: {email_match}\n')
                                elif phone_match:
                                    f.write(f'Phone: {phone_match}\n')


# change it to the output file.txt that was created on the last function
with open('path/to/the/outputfile/file.txt', 'r') as f:
    lines = f.readlines()

# Create an empty list to store the extracted data
data = []

# Loop over the lines in the output file
for line in lines:
    # Check if the line contains an email address or a phone number
    if 'Email' in line:
        # Extract the email address and append to data list as a new dictionary
        email = line.strip().split(': ')[1]
        data.append({'Email': email, 'Phone': None})
    elif 'Phone' in line:
        # Extract the phone number and update the last dictionary in the data list
        phone = line.strip().split(': ')[1]
        if data:
            data[-1]['Phone'] = phone

# Create a pandas DataFrame from the data list
df = pd.DataFrame(data)

# Save DataFrame to Excel file choose a path for saving the excel sheet
df.to_excel(r'path\to\save\the\excel\file\output.xlsx', index=False
