##Project: ICL02.txt
##Author: Elijah Allotey
##Date: 09/24/22
##Purpose: Grabs emails from user credentials and 
## formats them to a pdf in its own file. Then it places those files in a directory
## finally, at the end of day, zips files and uploads themtochosen shared GoogleDrive File




from __future__ import print_function
import email
import imaplib
import os
import os.path
import zipfile
import datetime
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from fpdf import FPDF
import fpdf
import time
import uuid
import shutil



fpdf.set_global("SYSTEM_TTFONTS", os.path.join(os.path.dirname(__file__),'fonts'))
SCOPES = ['https://www.googleapis.com/auth/drive.metadata.readonly', 'https://www.googleapis.com/auth/docs', 'https://www.googleapis.com/auth/drive.file']
ide = str(uuid.uuid1().hex[:8])
## Getting User Creds
usr = 'leatestm@gmail.com'
pasw = 'rnbduxvlljdlqelo' # This is not the users regular password, this is a special app password
imap_url = 'imap.gmail.com'

origin = '.'

if 'Mail' in os.listdir(origin):
    shutil.rmtree("Mail")
if 'Mail' not in os.listdir(origin):
    os.mkdir('Mail')

    

def ftext(mail):
    if mail.is_multi():
        return ftext(mail.get_payload(0))
    else:
        return mail.fetch_text(None, True)


#
def search(value, key, con):
	result, data = con.search(None, key, '"{}"'.format(value))
	return data
#


def fdata(bytes):
    mails = []
    for num in bytes[0].split():
        typ, data = ads.fetch(num, '(RFC822)')
        mails.append(data)
        
        return mails
    
    


          
def fetch(): 
    ads = imaplib.IMAP4_SSL(imap_url)
    ads.login(usr, pasw) # Logging in with creds
    ads.select('inbox') # Can be changed based on label
    result, data = ads.uid('search', None, "UNSEEN") 
    if result == 'OK':
            for num in data[0].split():
                result, data = ads.uid('fetch', num, '(RFC822)')
                if result == 'OK':
                    email_bytes = data[0][1]
                    email_str = email_bytes.decode('utf-8')
                    email_message = email.message_from_string(email_str)
                    From = ('From: ' + email_message['From'])
                    to = ('To: ' + email_message['To'])
                    date = ('Date: ' + email_message['Date'])
                    subj = ('Subject: ' + str(email_message['Subject']))
                    name = str(email_message['Subject'])
                    destination = "C:/Users/lelij/Work/Mail"
                    mail_dest = os.path.join(destination, name)
                    if not os.path.exists(mail_dest):
                        os.mkdir(mail_dest)
                    mail_dest_attach = os.path.join(mail_dest, 'attachments')
                    if not os.path.exists(mail_dest_attach):
                        os.mkdir(mail_dest_attach)
                               
                    for part in email_message.walk():
                        if part.get_content_type() == "text/plain":
                            body = part.get_payload(decode=True)
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get('Content-Disposition') is None:
                            continue
                        file_name = part.get_filename()
                        if bool(file_name):
                            pathto = os.path.join(origin, 'Mail', name, file_name,)
                            if not os.path.isfile(pathto) :
                                if os.path.exists(mail_dest_attach):
                                    
                                    move = open(pathto, 'wb')
                                    move.write(part.get_payload(decode=True))
                                    move.close()
                    
                    
                    mail_path = os.path.join(origin, 'Mail', mail_dest, name)
                    
                    mailwrite = FPDF()
                    mailwrite.add_font("NotoSans", style="", fname="NotoSans-Regular.ttf", uni=True)
                    mailwrite.add_page()
                    mailwrite.set_font("NotoSans", size = 12)
                    ## instead of writinng to a text file instead create the pdf and write to it before a txt file gets involved
                    
                    
                    
                    
                        
                    x = From
                    mailwrite.multi_cell(200, 10, txt = x, align = 'L')
                            
                        
                    x = to
                    mailwrite.multi_cell(200, 10, txt = x, align = 'L')
                            
                        
                    x = date
                    mailwrite.multi_cell(200, 10, txt = x, align = 'L')
                            
                        
                    x = subj
                    mailwrite.multi_cell(200, 10, txt = x, align = 'L')
                        
                    x = body.decode('utf-8')
                    mailwrite.multi_cell(200, 10, txt = x, align = 'L')
                    
                    mailwrite.output(mail_path + ".pdf")
                    
                    
                    
                    
                    
                    

def upload():
    #Getting Scope Permissions
    creds = None
    if os.path.exists('token.json'): 
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
    
        # Saves Permissions for future use
        with open('token.json', 'w') as permissions:
            permissions.write(creds.to_json())

    uplink = build('drive', 'v3', credentials=creds)
    now = datetime.now()
    date_time = now.strftime("%m_%d_%Y")
    metadata = {
        'name': 'Mail_' + date_time + "_" + usr + '.zip',
        'parents': ['1XpidUIPglE7iX6uX2NyL07XauNP95Kno']} #Denotes file upload location
    media = MediaFileUpload('Mail_' + date_time + "_" + usr +".zip",
                            mimetype='application/zip')
        
    uplink.files().create(body=metadata, media_body=media,
                                  fields='id', supportsAllDrives=True).execute()
        


def zip():
    now = datetime.now()
    date_time = now.strftime("%m_%d_%Y")
    

    zf = zipfile.ZipFile("Mail_" + date_time + "_"+ usr + ".zip", "w")
    for dirname, subdirs, files in os.walk("Mail"): 
        zf.write(dirname)
        for filename in files:
            zf.write(os.path.join(dirname, filename))
    zf.close()
    print("File Zipped")
    #old_name = r"C:\Users\lelij\Work\Mail.zip"
    #new_name = r"C:\Users\lelij\Work\Mail_" + date_time +".zip"

    #if os.path.isfile(new_name):
        
        #print("The file already exists")
    #else:
        #os.rename(old_name, new_name)


    
    
    
    
startup = 1
zipfunc = 0
nexti = 0
cou = 1

# Measuring elapsed time while collecting emails during the day till roughly 5:00
fetch()
zip()
upload()


            
               
               
        

                        

   

    


