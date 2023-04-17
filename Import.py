#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import os
import pandas as pd
import win32com.client as win32
import datetime


DEV_EMAILS = ["Rayon.Lewis@HealthTrustWS.com", "colin.jackson@healthtrustws.com"]

SHEET_NAMES = ['Air Data', 'Hotel Data', 'Car Data']
# Constants
EMAIL_SUBJECT = "Your HCA 26624 Daily Detail Report"
FILE_NAME = "Reservation Detail Super Summary"


EMAIL_SUBJECTS = ["Your HCA 26624 Daily Detail Report", "Your HCA 26624 Monthly Detail Report"]


ATTACHMENT_FOLDER = r"\\aastmgtmgt1\StaRN\BCD_Travel\test"
ATTACHMENT_FOLDER_TEST = r"C:\bcd_folder"
TEMPLATE_FILE = r"C:\template\template.xlsx"  # Replace with the path to your template file

global COUNTER
COUNTER = 0




# Check all sheets instead
SHEET_NAME = "Hotel"
SHEET_NAMES = ["Air Data", "Car Data", "Hotel Data"]


def process_attachments(attachments = None, time_string = None):

    global COUNTER
    for attachment in attachments:
        if attachment.FileName.endswith(".xlsx") and attachment.FileName.startswith("Reservation Detail Super Summary"):
            attachment.SaveAsFile(os.path.join(ATTACHMENT_FOLDER_TEST, attachment.FileName))

            os.rename(os.path.join(ATTACHMENT_FOLDER_TEST, attachment.FileName), os.path.join(ATTACHMENT_FOLDER_TEST, attachment.FileName.replace(".xlsx","{}.xlsx".format(time_string))) )

            #add_missing_headers(os.path.join(ATTACHMENT_FOLDER_TEST, attachment.FileName))
        




def add_missing_headers(file):
    
    for sheet in SHEET_NAMES:
        file_df= pd.read_excel(file, sheet_name=sheet)
        
        templ_df= pd.read_excel(TEMPLATE_FILE, sheet_name=sheet)
        
        ##compare the column sets for the template and each file
        if set(file_df.columns)!= set(templ_df.columns): 
            print(sheet, ' is not the same in ', file, "now fixed!")
            #For each sheet that does not have data this would be replaced with the template sheet
            with pd.ExcelWriter(file, mode = 'a', engine = 'openpyxl', if_sheet_exists='replace') as writer:
                templ_df.to_excel(writer, sheet_name = sheet, index = False)
        else:
            pass
        


def main():

    # Initialize Outlook Application and Namespace objects
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Specify the mailbox email address
    mailbox_email = "HWSBCDConcurReporting@HealthTrustWS.com"

    # Get the Recipient object for the specified mailbox
    recipient = outlook.CreateRecipient(mailbox_email)
    recipient.Resolve()

    # Get the inbox folder of the specified mailbox (6 refers to the Inbox folder)
    inbox = outlook.GetSharedDefaultFolder(recipient, 6)

    messages = inbox.Items.Restrict("@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{}%'".format(EMAIL_SUBJECT))


    print(len(messages))



    for message in messages:
        received_time = message.ReceivedTime
        received_time_str = received_time.strftime('%Y-%m-%d_%H-%M-%S')
        process_attachments(message.Attachments, received_time_str)

    print(f"Total Files: {len(messages)}; Corrected Files {COUNTER}")



def main2():
    os.chdir(r'C:\bcd_folder')
    files = [i for i in os.listdir() if i.startswith('Res')]
    
    for file in files:
        add_missing_headers(file)

if __name__ == "__main__":


    #add_missing_headers()
    main2()

