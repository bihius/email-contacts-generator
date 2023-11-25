################################################################################
#                                 description                                  #
# This script is used to generate a CSV file containing email addresses and    #
# names from all emails within an account.                                     #
#                                                                              #
# Created by: bihius                                                           #
################################################################################

import imaplib
import email
from email.utils import parseaddr
import csv
from tqdm import tqdm
import re
import os
import config 
from dotenv import load_dotenv # load environment variables from .env file for testing purposes

load_dotenv()
SERVER = os.getenv("SERVER") or config.SERVER
PORT = int(os.getenv("PORT") or config.PORT)  
EMAIL_USER = os.getenv("EMAIL_USER") or config.EMAIL_USER
EMAIL_PASS = os.getenv("EMAIL_PASS") or config.EMAIL_PASS
CSV_FILENAME = "contacts.csv"  

# constants
MAX_PROCESSED_MESSAGES = 10
CSV_HEADER = ["Name", "Email"]

# connect to email server
mail = imaplib.IMAP4_SSL(SERVER, PORT) if PORT == 993 else imaplib.IMAP4(SERVER, PORT)

def clean_name(name):
    # using decode_header() to decode name
    decoded_name, encoding = email.header.decode_header(name)[0]

    if isinstance(decoded_name, bytes):
        # if name contains bytes, decode them to utf-8
        try:
            name = decoded_name.decode(encoding or 'utf-8', errors='replace')
        except UnicodeDecodeError:
            name = "Invalid Encoding"
    else:
        name = str(decoded_name)

    # remove trashy characters
    name = re.sub(r'[^a-zA-Z0-9\s]', '', name)
    name = name.strip()

    return name

try:
    # login
    mail.login(EMAIL_USER, EMAIL_PASS)

    # initialize variable for output
    all_addresses = set()

    # for testing purposes only
    processed_messages = 0

    # get list of all folders
    status, folders = mail.list()

    if status == "OK":
        # loop through all folders
        for folder in folders:
            # get folder name
            folder_name = folder.split()[-1].decode('utf-8')

            # select folder name
            status, messages = mail.select(folder_name)
            if status == "OK":
                # search for all messages in folder
                status, messages = mail.search(None, "ALL")
                messages = messages[0].split()

                # loop through all messages in folder
                for mail_id in tqdm(messages, desc=f"Processing '{folder_name}'", total=len(messages), position=0, bar_format='{l_bar}{bar}'):
                    if processed_messages >= MAX_PROCESSED_MESSAGES:
                        break

                    try:
                        _, msg_data = mail.fetch(mail_id, "(RFC822)")
                        msg = email.message_from_bytes(msg_data[0][1])

                        # get recipient's name and email address
                        from_name, from_address = parseaddr(msg.get("From"))

                        # clean recipient's name
                        from_name = clean_name(from_name)

                        # add recipient's name and email address to data variable
                        all_addresses.add((from_name, from_address.strip()))

                        # limit processing count to 10 messages, for testing purposes only
                        # processed_messages += 1

                    except Exception as e:
                        print(f"Error processing message {mail_id}: {e}")

    # close connection to email server
    mail.logout()

    # write to CSV file
    with open(CSV_FILENAME, mode="w", newline="", encoding="utf-8") as csvfile:
        # create csv writer object
        csv_writer = csv.writer(csvfile)

        # write header row
        csv_writer.writerow(CSV_HEADER)
        
        #loop through all contacts in data variable
        for from_name, from_address in all_addresses:
            csv_writer.writerow([from_name, from_address])

    print(f"Your contacts have been saved to {CSV_FILENAME}.")

except Exception as e:
    print(f"Error: {e}")

finally:
    # close CSV file
    csvfile.close()
