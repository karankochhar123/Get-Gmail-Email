import email
import imaplib
import sys
import pandas as pd

#Get Username and Password
Username = input("Enter User Name -->")
Password = input("Enter Password --->")


#Please ensure that gmail security access has been disbled for less secure apps
#link for how to diable access https://support.google.com/cloudidentity/answer/6260879?hl=en

#Connecting to gmail

try:
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(Username, Password)
    print ("Connection Succesful")
except:
    sys.exit(1)


#Getting user input for email date ranges 
mail.select("inbox")
From_Date = input("Enter From date in DD-MMM-YYYY format--->")
To_Date = input("Enter To date in DD-MMM-YYYY format------->")


#Going through emails  
typ, data = mail.search(None,'(SINCE "%s")' %str(From_Date),'(BEFORE "%s")' %str(To_Date))

inbox_item_list = data[0].split()

Totalemails=len(inbox_item_list)

print("Total Emails ---->"+str(Totalemails))

lst = []

for item in inbox_item_list:
    typ2, email_data = mail.fetch(item,'(RFC822)')
    for response_part in email_data:
        if isinstance(response_part, tuple):
            email_message = email.message_from_string(str(response_part[1].decode("utf-8","ignore")))
            Date_ = email_message['Date']
            From_ = email_message['From']
            To_ = email_message['To']
            Sub_ = email_message['Subject']
            print("Processing email-->"+str(Totalemails))
            lst.append([Date_,From_,To_,Sub_])
            Totalemails=Totalemails-1
            

#saving emails to CSV file
cols = ['Date','From','To','Subject']
df = pd.DataFrame(lst,columns=cols)
df.to_csv('Emails.csv')
print("File Saved")


