# Download payslips from your inbox
# (c) Najeem Muhammed

import win32com.client
import os.path

def get_items_recurse(folder):
    print("Checking in folder: {}".format(folder.name))
    items = []
    items.append(folder.items)
    for f in folder.folders:
        for item in get_items_recurse(f):
            items.append(item)
    return items

path = input("Enter the path where you want to dump the files: ")
if not os.path.exists(path):
    input("Path doesn't exist! Press Enter to exit.")
    exit()
searchtext = "payslip"
print("Downloading all pdf files from mails with Payslip in subject......")
try:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
except:
    raise("Make sure outlook is open when running this code!")
inbox = outlook.GetDefaultFolder(6)
messages_list = get_items_recurse(inbox)
for messages in messages_list:
    for msg in messages:
        if searchtext.lower() in msg.subject.lower():
            if msg.attachments.count:
                filename = msg.attachments.item(1).filename
                if '.pdf' in filename.lower():
                    print(msg.subject)
                    msg.attachments.item(1).saveasfile(os.path.join(path, filename))
                else:
                    print("{} not saved".format(filename))