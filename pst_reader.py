import base64

your_code = base64.b64encode(b"""

from pathlib import Path  #core python module
import win32com.client  #pip install pywin32
import mimetypes
import os
import datetime
import time
import csv
import random
import string
import sys
import zipfile
import stat
import shutil
import glob
import json
from time import sleep

def zipdir(path, ziph):
    # ziph is zipfile handle
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(root, file),
                       os.path.relpath(os.path.join(root, file),
                                       os.path.join(path, '..')))

def find_pst_folder(OutlookObj, pst_filepath) :
    for Store in OutlookObj.Stores :
        if Store.IsDataFileStore and Store.FilePath == pst_filepath :
            return Store.GetRootFolder()
    return None

def enumerate_folders(FolderObj) :
    for ChildFolder in FolderObj.Folders :
        enumerate_folders(ChildFolder)

    if str(FolderObj) in folder_list or "(This computer only)" in str(FolderObj):
        return
    iterate_messages(FolderObj)

def iterate_messages(FolderObj) :
    print("Reading folder => ", FolderObj)
    count = 0
    for index, message in enumerate(FolderObj.Items):
        check = message.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x001A001E")
        if check == "IPM.contact" or check == "IPM.Appointment":
            break
        if check != "IPM.Note":
            continue

        subject = message.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x0037001E")
        if subject == "Synchronization Log:":
            continue
        sender_name = message.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1A001E")
        if sender_name == "Microsoft Outlook" or sender_name == "Antispam":
            continue
        sender_email = message.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001F")
        if sender_name == ";" or sender_name == "":
            sender_name = "unknown"
        if sender_email == ";" or sender_email == "":
            sender_email = "unknown"

        try:
            recipients = message.Recipients
            recipient_to_name = []
            recipient_to_address = []
            recipient_cc_name = []
            recipient_cc_address = []
            for recipient in recipients:
                try:
                    if recipient.Type == 1:
                        if recipient.Name == ";" or recipient.Name == "":
                            recipient_to_name.append("unknown")
                        else:
                            recipient_to_name.append(recipient.Name.replace(",", ""))

                        if recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress == ";" or recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress == "":
                            recipient_to_address.append("unknown")
                        else:
                            recipient_to_address.append(recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress)
                    elif recipient.Type == 2:
                        if recipient.Name == ";" or recipient.Name == "":
                            recipient_cc_name.append("unknown")
                        else:
                            recipient_cc_name.append(recipient.Name.replace(",", ""))
                        if recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress == ";" or recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress == "":
                            recipient_cc_address.append("unknown")
                        else:
                            recipient_cc_address.append(recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress)
                except:
                    if recipient.Type == 1:
                        if recipient.Name == ";" or recipient.Name == "":
                            recipient_to_name.append("unknown")
                        else:
                            recipient_to_name.append(recipient.Name.replace(",", ""))
                        if recipient.Address == ";" or recipient.Address == "":
                            recipient_to_address.append("unknown")
                        else:
                            recipient_to_address.append(recipient.Address)
                    elif recipient.Type == 2:
                        if recipient.Name == ";" or recipient.Name == "":
                            recipient_cc_name.append("unknown")
                        else:
                            recipient_cc_name.append(recipient.Name.replace(",", ""))
                        if recipient.Address == ";" or recipient.Address == "":
                            recipient_cc_address.append("unknown")
                        else:
                            recipient_cc_address.append(recipient.Address)

            recipient_to_name_str = ','.join(recipient_to_name)
            recipient_to_address_str = ','.join(recipient_to_address)
            recipient_cc_name_str = ','.join(recipient_cc_name)
            recipient_cc_address_str = ','.join(recipient_cc_address)
        except:
            recipient_to_name_str = ""
            recipient_to_address_str = ""
            recipient_cc_name_str = ""
            recipient_cc_address_str = ""


        # if subject == "" or (sender_name == "" or recipient_to_name_str == ""):
        #     continue
        date = message.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x00390040")
        receive_date = ""
        if date is not None:
            receive_date = message.PropertyAccessor.UTCToLocalTime(date)
        body = message.HTMLBody
        try:
            attachments = message.Attachments
            folder = ""
            attachment_array = []
            attachment_str = ""
            if len(attachments) > 0:
                folder = f"{str(time.time())}_{''.join(random.choice(string.ascii_lowercase) for i in range(5))}"
                target_folder = output_dir / folder
                os.makedirs(target_folder,777,True)

                # Save attachments
                for attachment in attachments:
                    # if mimetypes.guess_type(str(attachment))[0] is not None:
                    attachment_array.append(attachment.Filename.replace(",", ""))
                    attachment.SaveAsFile(target_folder / attachment.Filename.replace(",", ""))

                attachment_str = ','.join(attachment_array)
        except:
            attachment_str = ""
            folder = ""
        rowdatas.append({
            'date' : str(receive_date),
            'subject' : subject,
            'from_name' : sender_name.replace(",", ""),
            'from_email' : sender_email,
            'to_name' : recipient_to_name_str,
            'to_email' : recipient_to_address_str,
            'cc_name' : recipient_cc_name_str,
            'cc_email' : recipient_cc_address_str,
            'body' : body,
            'attachment' : attachment_str,
            'attachment_folder' : folder,
            'folder' : str(FolderObj),
            'pst' : file
        })
        count +=1
        print("Extracted ", count ," email(s) from ", file, "folder =>", FolderObj)
    print("Ended folder => ", FolderObj)

folder_list = ['Recipient Cache', 'Sync Issues']
output_dir = Path.cwd() / "pst"
pst_dir = Path.cwd() / "put-your-pst-here"
os.makedirs(output_dir,777,True)
os.chmod(output_dir, 0o0777)
rowdatas = []
pst = ""
size = 0
fileNumber = 0
os.chdir(pst_dir)
for file in glob.glob("*.pst"):
    fp = os.path.join(pst_dir, file)
    shutil.copy(fp, output_dir)
    size += os.path.getsize(fp)
    fileNumber += 1
if size/1024/1024 > 1024 :
    print("Files' size exceed limit of 1GB!")
    sleep(0.50)
    input("Press Enter to quit...")
    exit()
if fileNumber == 0 :
    print("No file is scanned")
    sleep(0.50)
    input("Press Enter to quit...")
    exit()
for file in glob.glob("*.pst"):
    pst = os.path.join(pst_dir, file)
    print("Reading file: ", file)
    try:
        Outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        Outlook.AddStore(pst)
        PSTFolderObj = find_pst_folder(Outlook,pst)
        try :
            print("Extracting file: ", file)
            enumerate_folders(PSTFolderObj)
            print("Done Extracting: ", file)
        except Exception as exc :
            print(exc)
            input("Press Enter to quit...")
            exit()
        finally:
            Outlook.RemoveStore(PSTFolderObj)

    except Exception as exc :
        print(exc)
        input("Press Enter to quit...")
        exit()

print("Storing emails content into json.")
sleep(0.50)

with open(output_dir / f"{''.join(random.choice(string.ascii_lowercase) for i in range(5))}.json", 'w', encoding='UTF8') as outfile:
    json.dump(rowdatas, outfile)
#
# header = ['date', 'subject', 'from_name', 'from_email', "to_name", "to_email", "cc_name", "cc_email", "body", "attachment", "attachment_folder","pst"]
# with open(output_dir / f"{''.join(random.choice(string.ascii_lowercase) for i in range(5))}.csv", 'w', encoding='UTF8') as f:
#     writer = csv.writer(f)
#     writer.writerow(header)
#     for rowdata in rowdatas:
#         writer.writerow(rowdata)
# # close the file
# f.close()
print("Saved ", len(rowdatas), "emails")
sleep(0.50)


try:
    print("Zipping folder, please wait");
    name = f"{''.join(random.choice(string.ascii_lowercase) for i in range(5))}.zip"
    with zipfile.ZipFile(name, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipdir(output_dir, zipf)

    print("Done! Please find your zipped file (",name,") in ", pst_dir)
    sleep(0.50)
except Exception as exc :
    print("Failed to zip. Please try again.")
    input("Press Enter to quit...")
    exit()
finally :
    os.system("taskkill /f /im Outlook.exe")
    try :
        shutil.rmtree(output_dir)
    except Exception as exc :
        print("Cannot remove folder.")
        input("Press Enter to quit...")
        exit()


input("Press Enter to quit...")

""")

exec(base64.b64decode(your_code))
