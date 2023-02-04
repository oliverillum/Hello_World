import os
import win32com.client as win32

# construct Outlook application instance

olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')

User = input("which user are running this script? ")
period = input("what is the period 'yyyy MM'?" )
customer = input("which customer do you want to send for?")

data = {
    'Reipurth Dentalservice': {'email': '<sl@timevat.com>', 'path': 'C:\\Users\\'+ User +'\\TIMEVAT A S\\Kommunikationswebsted - TIMEVAT\\Operation\\Reipurth Dentalservice'},
    'Soft Sales': {'email': 'ob@timevat.com', 'path': 'C:\\Users\\'+ User +'\\TIMEVAT A S\\Kommunikationswebsted - TIMEVAT\Operation\\Soft Sales'},
    'Skall Studio': {'email': 'vl@timevat.com', 'path': 'C:\\Users\\'+ User +'\\TIMEVAT A S\\Kommunikationswebsted - TIMEVAT\Operation\\Skall Studio ApS'}
}

pathToCustomerFolder = data[customer]['path']
mailForCustomer = data[customer]['email']

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


attachment = ((path) + "\\" + (target_name))



# construct the email item object
mailItem = olApp.CreateItem(0)
mailItem.Subject = ((target_name[:-4]) + " " + customer)
mailItem.BodyFormat = 1
mailItem.Body = "Hello World"
mailItem.To = (mailForCustomer)
mailItem.Attachments.Add(attachment)



# mailItem.Display()/mailItem.Send()/mailItem.Save()

mailItem.Display()
