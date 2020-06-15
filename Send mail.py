from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from email.mime.text import MIMEText
from apiclient import errors
import random
import string
import base64
import os.path
import pickle
from xlrd import open_workbook
from xlwt import Workbook


def getrand(l=8):
    key = string.digits + string.ascii_letters + string.digits
    return ''.join((random.choice(key) for i in range(l)))


def loadcreds():
    # If modifying these scopes, delete the file token.pickle.
    SCOPES = ['https://www.googleapis.com/auth/gmail.send',
              'https://www.googleapis.com/auth/gmail.compose']

    creds = None

    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds


def SendMessage(service, user_id, message):
    try:
        message = (service.users().messages().send(userId=user_id, body=message)
                   .execute())
        print('Success! Message Id: %s' % message['id'])
        return True
    except errors.HttpError as error:
        print('An error occurred: %s' % error)
        return False


def CreateMessage(email, vid):
    link = 'https://montagekmc.typeform.com/to/U7Mbbh'
    test='<html><br><h1><b>This is a Test Email!</b></h1>You <b>will</b> receive a email from <i>this</i>' \
         ' <span style="color: cornflowerblue">(montagevoting@gmail.com)</span> email id with some instructions,' \
         ' please read them carefully before start voting.<br><br> Your Unique Voting ID will be something like this:' \
         ' <b>{}</b><br><span style="font-size: 20">THIS IS <b style="color: crimson">NOT</b> YOUR OFFICIAL ID</span>' \
         '<br><br>To Prevent the email from going in your spam or any other label, just reply to this email with a message' \
         ' like <b style="color: darkgreen">"Hi"</b><br><br>Also check if <a href="montagekmc.media/vote" style="color: ' \
         '#21ce99">links</a> work properly, they should redirect you to the page where the voting will take place. ' \
         'You will recieve the final voting link in the email itself.<br><br>Thanks<br>Kapil</html>'.format(vid)
    messagehtml = '<html><body><font face="arial, sans-serif">Here is your Unique&nbsp;</font>Voter&nbsp;<font face="arial,' \
                  ' sans-serif">ID: <b>{}</b><br><i>Required to register your vote</i></font><div><br></div><div><b>' \
                  '<font color="#ff0000">NOTE:</font></b></div><div><font style="" color="#000000">1.&nbsp;</font>This ' \
                  'key is one time use only and cannot be generated again, so<b> Don\'t lose it!</b></div><div>2.&nbsp;Any' \
                  ' double entries will result in your vote being ineligible for counting, so <b>vote carefully!</b>' \
                  '</div><div>3.&nbsp;If the ID' \
                  ' doesn\'t match <i>exactly</i>, your vote won\'t be registered, so <b>Copy-Paste</b> it.</div><div>4. ' \
                  'You have 24 hrs to vote, ie. Voting ends <b>20:00 12 June 2020.</b></div><div>5. Read and fill the form carefully, if you face any difficulty or have a query just drop a message.</div><div><br></div><div>' \
                  '<font size="4">You can go and vote using this <a style="color: #21ce99; text-decoration: none" ' \
                  'href="{}"><b>link</b></a></font></div></body></html>'.format(vid, link)
    message = MIMEText(messagehtml, 'html')
    message['to'] = email
    message['from'] = 'Team Montage'
    message['subject'] = 'Montage Elections'
    return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}


def reademailids(path):
    sheet = open_workbook(path).sheet_by_index(0)
    [email_at] = [i for i, x in enumerate(sheet.row_values(0)) if x == 'Email']
    return [x.value for x in sheet.col(email_at) if x.ctype is not 0][1:]

def writesheet(lst,name):
    wb=Workbook()
    sheet=wb.add_sheet('Sheet 1')
    for i,vid in enumerate(lst):
        sheet.write(i,0,vid)
    wb.save(name)

## Driver Code:

service = build('gmail', 'v1', credentials=loadcreds())
emails = reademailids('Electoral Roll.xlsx')
nemails=len(emails)
counter=1
vids=[]
fmails=[]
for to in emails:
    vid = getrand()
    message = CreateMessage(to, vid)
    sent = SendMessage(service, "me", message)
    if sent: vids.append(vid)
    else:
        fmails.append(to)
        print("Failed: ",to)
        counter+=1

writesheet(vids,'Vids.xls')

if fmails:
    print('Couldn\'t send to following emails:',fmails)
    writesheet(fmails,'Failed Mails.xls')
else:
    print("{} Emails successfully sent.".format(len(emails)))
