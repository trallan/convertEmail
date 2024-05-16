import win32com.client
import csv
import os
import re
from datetime import datetime, timedelta
# outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# msg = outlook.OpenSharedItem(r"C:\Users\SveaUser\Desktop\ConvertEmail\test.msg")

msg_directory = r"C:\Users\SveaUser\Desktop\ConvertEmail\emails"
msg_files = [file for file in os.listdir(msg_directory) if file.endswith(".msg")]
phone_number_pattern = r"\b(?:\+\d{1,2}\s)?\(?(?:\d{3}[-.\s]?\d{7}|\d{3}[-.\s]?\d{3}[-.\s]?\d{4})\b"

with open('output.csv', 'w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(['name', 'date/time', 'phonenumber1', 'phonenumber2', 'Framgår ordet Svea'])
    for msg_file in msg_files:
         # Construct the full path to the .msg file
        msg_path = os.path.join(msg_directory, msg_file)
        
        # Open the .msg file
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        msg = outlook.OpenSharedItem(msg_path)
        
        sender_name = msg.SenderName
        sent_date_time = msg.Senton - timedelta(hours=2)

        phone_numbers = re.findall(phone_number_pattern, msg.Body)

        lower_body = msg.Body.lower()

        contains_svea = any(keyword in lower_body for keyword in ['svea inkasso', 'sveainkasso', 'svea bank', 'sveabank'])
        
        if len(phone_numbers) > 1:
            writer.writerow([sender_name, sent_date_time, phone_numbers[0], phone_numbers[1], contains_svea])
        elif len(phone_numbers) == 1:
            writer.writerow([sender_name, sent_date_time, phone_numbers[0], '', contains_svea])
        else:
            writer.writerow([sender_name, sent_date_time, '', '', contains_svea])

        # Closing the message object after you're done with it
        del outlook, msg

# print(msg.SenderName)
# print(msg.SenderEmailAddress)
# print(msg.SentOn)
# print(msg.To)
# print(msg.CC)
# print(msg.BCC)
# print(msg.Subject)
# print(msg.Body)

# count_attachments = msg.Attachments.Count
# if count_attachments > 0:
#     for item in range(count_attachments):
#         print(msg.Attachments.Item(item + 1).Filename)

# del outlook, msg