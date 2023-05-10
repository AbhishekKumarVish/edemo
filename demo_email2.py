import os
import pandas as pd
from email import policy
from email.parser import BytesParser
import chardet

# Directory containing .eml files
directory = r'F:\srk\Edemo'

# Initialize an empty list to store email data
emails_data = []

# Recursive function to extract email body
def extract_body(part):
    if part.is_multipart():
        return [extract_body(subpart) for subpart in part.get_payload()]
    else:
        try:
            # Attempt to decode the body using UTF-8 encoding
            return part.get_payload(decode=True).decode('utf-8')
        except UnicodeDecodeError:
            # If decoding with UTF-8 fails, try detecting the encoding and decode accordingly
            encoding = chardet.detect(part.get_payload())['encoding']
            return part.get_payload(decode=True).decode(encoding, errors='replace')

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
            body = extract_body(msg)

            # Append the email data to the list
            emails_data.append({
                'Subject': subject,
                'From': from_address,
                'To': to_address,
                'Date': date,
                'Body': body
            })

# Convert the list of dictionaries into a dataframe
edf = pd.DataFrame(emails_data)
