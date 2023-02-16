import os
import win32com.client as win32
import logging
import pandas as pd
import base64
import UserData

User = UserData.username
Name2 = UserData.name1
TimevatMail = UserData.Mail

#User = input("which user are running this script? ")
#User = "olive"
#Name2 = "Oliver Bregneberg"
#TimevatMail = "ob@timevat.com"

# set up logging
logging.basicConfig(filename='email_script.log', level=logging.DEBUG)

# construct Outlook application instance
olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')

df = pd.read_excel("C:\\Users\\" + User + "\\TIMEVAT A S\\Kommunikationswebsted - TIMEVAT\\IT og Projekter\\Python\\Tolddeklarationsoversigt sender\\clients.xlsx")

# constants
SUBJECT = "Tolddeklarationsoversigt {} ({})"
ATTACHMENT_NAME = "{} Tolddeklarationsoversigt.pdf"
BODY = "Template2"

#period = input("What is the period of the declarations? XXXX XX ")
period = "2023 01"

#customer = input("VAT nr. for the customer you want to send to? ")
customer = "913397045"
df['VAT'] = df['VAT'].astype(int)

customer = int(customer)

pathToCustomerFolder = df.loc[df['VAT'] == customer, 'Folder'].values[0] 
nameToCustomer = df.loc[df['VAT'] == customer, 'Name'].values[0]
mailForCustomer = df.loc[df['VAT'] == customer, 'Email_to'].values[0]

ROOT_DIR = "C:\\Users\\" + User + "\\TIMEVAT A S\\Kommunikationswebsted - TIMEVAT\\Operation\\" + nameToCustomer

def find_path(name, path):
    for root, dirs, files in os.walk(path):
        if name in dirs or name in files:
            return root
    return None

target_name = (period +" Tolddeklarationsoversigt.pdf")
path = find_path(target_name, ROOT_DIR)
if path:
    print(f"Found at: {path}")
    logging.info(f"{target_name} was successfully found " f"({User} {period} {nameToCustomer})")
    
else:
    print(f"{target_name} not found")
    logging.error(f"{target_name} was not found " f"({User} {period} {nameToCustomer})")

attachment = ((path) + "\\" + (target_name))


# construct the email item object
mailItem = olApp.CreateItem(0)
mailItem.Subject = ((target_name[:-4]) + " " + nameToCustomer)
mailItem.BodyFormat = 1

with open("C:\\Users\\olive\\TIMEVAT A S\\Kommunikationswebsted - TIMEVAT\\IT og Projekter\\Python\\Tolddeklarationsoversigt sender\\TimevatLogo.png", "rb") as image_file:
    image_data = image_file.read()

# Encode the image data as a base64 string
encoded_image = base64.b64encode(image_data).decode("utf-8")

with open("C:\\Users\\olive\\TIMEVAT A S\\Kommunikationswebsted - TIMEVAT\\IT og Projekter\\Python\\Tolddeklarationsoversigt sender\\TimeVatWebsiteImg.png", "rb") as image_fileWebsite:
    image_dataWebsite = image_fileWebsite.read()

# Encode the image data as a base64 string
encoded_imageWebsite = base64.b64encode(image_dataWebsite).decode("utf-8")

with open("C:\\Users\\olive\\TIMEVAT A S\\Kommunikationswebsted - TIMEVAT\\IT og Projekter\\Python\\Tolddeklarationsoversigt sender\\TimeVatLinkdImg.png", "rb") as image_fileLinkdIn:
    image_dataLinkdIn = image_fileLinkdIn.read()

# Encode the image data as a base64 string
encoded_imageLinkdIn = base64.b64encode(image_dataLinkdIn).decode("utf-8")

mailItem.htmlBody = f"""
<html>
  <body style="font-family: Calibri Light, Arial, sans-serif; font-size: 16px;">
    <div style="padding: 20px; color: #333;">

      <p style="margin-bottom: 10px;">Hej,</p>
      <p>Vedhæftet finder du kopi af Tolddeklarationsoversigt for {period}.</p>
      <p style="margin-bottom: 10px;">Med venlig hilsen / Best regards,</p>

      <div style="margin-top: 10px; line-height: 1.5;">
        <p>{Name2}<br>VAT Operation</p>
        
         <table style="font-family: Calibri Light, Arial, sans-serif; font-size: 16px;">
          <tr>
            <td>Phone:</td>
            <td>+45 7021 3000</td>
          </tr>
          <tr>
            <td>E-mail:</td>
            <td><a href="mailto: {TimevatMail}" style="color: blue; text-decoration: underline;">ob@timevat.com</a></td>
          </tr>
          <tr>
            <td>Web:</td>
            <td><a href="https://www.timevat.com" style="color: blue; text-decoration: underline;">www.timevat.com</a></td>
          </tr>
          <tr>
            <td>Address:</td>
            <td><u>Høje Taastrup Boulevard 52, DK-2630 Taastrup</u></td>
          </tr>
          <tr>
            <td>CVR:</td>
            <td>40 68 87 65</td>
          </tr>
        </table>
        <br>
        <div style="height: 20px;"></div>
        <img src="data:image/jpeg;base64,{encoded_image}" width="250" height="45" style="margin-bottom: -10px;">
        <p style="font-family: Calibri, sans-serif; margin-top: 1px;">Your foreign finance and export specialist.</p>
    <table>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      <td>
        <a href="https://twitter.com/your-twitter-handle-here">
          <img src="data:image/jpeg;base64,{encoded_imageLinkdIn}" width="50" height="50">
        </a>
      </td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      <td>
        <a href="https://www.timevat.com/">
          <img src="data:image/jpeg;base64,{encoded_imageWebsite}" width="50" height="50">
        </a>
      </td>
    </tr>
  </table>
      </div>
    </div>
  </body>
</html>

""".format(period)

mailItem.To = (mailForCustomer)
mailItem.Attachments.Add(attachment)
#mailItem.CC = input()

###

#choice = input("Do you want to display or send the email? (Send = 1, Display = 2) ")
#if choice.lower() == "2":
#    mailItem.Display()
#elif choice.lower() == "1":
#    mailItem.Send()
#else:
#    print("Invalid choice. Try again.")

mailItem.Display()
