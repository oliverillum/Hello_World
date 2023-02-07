import os
import win32com.client as win32
import logging
import pandas as pd


# set up logging
logging.basicConfig(filename='email_script.log', level=logging.DEBUG)

# construct Outlook application instance
olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')

df = pd.read_excel("clients.xlsx")


# constants
SUBJECT = "Tolddeklarationsoversigt {} ({})"
ATTACHMENT_NAME = "{} Tolddeklarationsoversigt.pdf"
ROOT_DIR = "C:\\Users\\Oliver\\Desktop\\\\TIMEVAT A S\\Kommunikationswebsted - TIMEVAT\\Operation"
BODY = "Template2"

#User = input("which user are running this script? ")
period = input("what is the period 'yyyy MM'?" )
customer = input("VAT nr. for the customer you want to send to?")
df['VAT'] = df['VAT'].astype(int)
customer = int(customer)

pathToCustomerFolder = df.loc[df['VAT'] == customer, 'Folder'].values[0] 
nameToCustomer = df.loc[df['VAT'] == customer, 'Name'].values[0]
mailForCustomer = df.loc[df['VAT'] == customer, 'Email_to'].values[0]

def find_path(name, path):
    for root, dirs, files in os.walk(path):
        if name in dirs or name in files:
            return root
    return None

root_dir = (pathToCustomerFolder)
target_name = (period +" Tolddeklarationsoversigt.pdf")
path = find_path(target_name, root_dir)
if path:
    print(f"Found at: {path}")
    
else:
    print(f"{target_name} not found")
    logging.error(f"{target_name} not found")

attachment = ((path) + "\\" + (target_name))





# construct the email item object
mailItem = olApp.CreateItem(0)
mailItem.Subject = ((target_name[:-4]) + " " + nameToCustomer)
mailItem.BodyFormat = 1
mailItem.Body = "Hello World"
mailItem.To = (mailForCustomer)
mailItem.Attachments.Add(attachment)
#mailItem.CC = input()



# mailItem.Display()/mailItem.Send()/mailItem.Save()

mailItem.Display()
