import base64
import os
import sys
from threading import Thread

import mimetypes
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

import json
import requests


EMAIL_LOG = {}
COUNTER = 0

SCOPES = ['https://mail.google.com/']

def error(msg):
    """It prints the Error and end the script."""
    print(msg)
    sys.exit(1)

def check_credential():
    """
      Checks the credentials and update it in GoogleApi/cred/token.json.
    """
    creds = None

    if os.path.exists('GoogleApi/cred/token.json'):
        creds = Credentials.from_authorized_user_file('GoogleApi/cred/token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            file_name = input("Enter the credentials.json file location : ")
            if os.path.exists(file_name):
                flow = InstalledAppFlow.from_client_secrets_file(
                    file_name, SCOPES)
                creds = flow.run_local_server(port=0)
            else:
                print("invalid location")

        with open('GoogleApi/cred/token.json', 'w') as token:
            token.write(creds.to_json())


def make_post_request(url,data=None,json=None,header=None):
    """
      Args
        url (str) : url endpoint where request to be made.
        data (dict) : path parameters (optional)
        json (dict) : request body parameters (optional)
        header (dict) : Header to be send sent in request (optional)

      Description:
        It sends a Post request to sepcified url endpoint and handels the
        Network connection error exceptions.
    """
    try:
        response=requests.post(url,data=data,json=json,headers=header)
        return response
    except requests.exceptions.ConnectionError:
        error("Kindly check Your internet Connection.")


def get_access_token():
    """
        Gives access token form GoogleApi/cred/token.json or creates one with the
        Deatails in credentials.json.
    """
    check_credential()
    with open('GoogleApi/cred/token.json') as token:
        raw_data = json.load(token)
        return raw_data['token']


def make_header():
    """ It makes a headers by passing the auth token and returns as dict"""
    header = {
      "Authorization" : f'Bearer {get_access_token()}'
    }
    return header


def create_message_with_attachment(
      to , subject,
    message_text , file
   ):
    """Create a message for an email.

    Args:
      sender: Email address of the sender.
      to: Email address of the receiver.
      subject: The subject of the email message.
      message_text: The text of the email message.
      file: The path to the file to be attached.

    Returns:
      An object containing a base64url encoded email object.
    """
    message = MIMEMultipart()
    message['to'] = to

    message['subject'] = subject

    msg = MIMEText(message_text)
    message.attach(msg)

    content_type, encoding = mimetypes.guess_type(file)

    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'

    main_type, sub_type = content_type.split('/', 1)
    fp = open(file, 'rb')
    msg = MIMEBase(main_type, sub_type)
    msg.set_payload(fp.read())
    fp.close()
    filename = os.path.basename(file)

    msg.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.add_header('Content-Type', main_type, name=filename)

    email.encoders.encode_base64(msg)
    message.attach(msg)

    return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}


def send_mail(to,subject,content,filename):
    """
      Args:
        to (str) : To address of the mail.
        subject (str) : subject of mail.
        filename(str) : file-location of attachment to be sent.

      Description:
        It sends email with attachments to participants.

    """
    global COUNTER

    msg_instance = create_message_with_attachment(to,subject,content,filename)
    header = make_header()
    url = f'https://gmail.googleapis.com/gmail/v1/users/me/messages/send'
    response = make_post_request(url,json=msg_instance,header=header)
    print(response.text)
  # print(response.status_code)
    COUNTER+=1
    if response.status_code != 200:
        EMAIL_LOG[len(EMAIL_LOG)+1] = [to,filename]
        COUNTER-=1
    # error("Error occured While Sending the mail.")



def send_mail_for_participants(data,subject,content):
    """
      Args:
        data (dict) : { key(int) : value(list) } value has to elemets namely,
                      mail-Id , attachement file path.
                      eg : {1: ["xxxx@gmail.com","root/home/dowloads/kkk.pdf"]}
      Description:
        It sends the mail to the participants in 15 Threads to complete the process
        as soon as possible.
    """
    global EMAIL_LOG
    EMAIL_LOG = {}
    index = 1
    number_of_threads = 10 # Defining the number of thread to be excuted.

    while index <=len(data):
        thread_list = []

        for i in range(number_of_threads):
            try:
                temp = Thread(target= send_mail,args =(
                data[index][0],
                subject,
                content,
                data[index][1]
              ))
            except KeyError:
                break
            temp.start()
            thread_list.append(temp)
            index+=1

        for thread in thread_list:
      # joining the threads with main thread.
            thread.join()

    if EMAIL_LOG:
        print("sending failed tries..")
        send_mail_for_participants(EMAIL_LOG)

    print(COUNTER)