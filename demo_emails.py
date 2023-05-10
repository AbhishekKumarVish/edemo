import os
import pandas as pd
from email import policy
from email.parser import BytesParser

# Directory containing .eml files
directory = r'F:\srk\Edemo'

# Initialize an empty list to store email data
emails_data = []

# Iterate over each .eml file in the directory
for filename in os.listdir(directory):
    if filename.endswith(".eml"):
        file_path = os.path.join(directory, filename)

        with open(file_path, 'rb') as fp:
            # Parse the .eml file
            try:
                msg = BytesParser(policy=policy.default).parse(fp)
            except UnicodeDecodeError:
                continue

            # Extract desired information from the email
            subject = msg['subject']
            from_address = msg['from']
            to_address = msg['to']
            date = msg['date']
            body = ''

            # Extract email body
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    if content_type == 'text/plain':
                        try:
                            body = part.get_payload(decode=True).decode('utf-8')
                        except UnicodeDecodeError:
                            continue
            else:
                try:
                    body = msg.get_payload(decode=True).decode('utf-8')
                except UnicodeDecodeError:
                    continue

            # Append the email data to the list
            emails_data.append({
                'Subject': subject,
                'From': from_address,
                'To': to_address,
                'Date': date,
                'Body': body
            })

# Convert the list of dictionaries into a dataframe
df = pd.DataFrame(emails_data)


df['Body']