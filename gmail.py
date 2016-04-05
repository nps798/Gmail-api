
from __future__ import print_function
import httplib2
import os

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools


from apiclient import errors
#上面這行是list message所引入的

import csv
#上面這行是為了要write dict 到 外面file裡面
try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/gmail-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/gmail.readonly'
CLIENT_SECRET_FILE = 'C:\\Users\\CHICHI\\Desktop\\client_secret_905212192142-kg04ho7bii3uo8sdv20mkom3rrll79uk.apps.googleusercontent.com.json'
APPLICATION_NAME = 'Gmail API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'gmail-python-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

"""Get a list of Messages from the user's mailbox.
"""




def ListMessagesMatchingQuery(service, user_id, query=''):
  """List all Messages of the user's mailbox matching the query.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    query: String used to filter messages returned.
    Eg.- 'from:user@some_domain.com' for Messages from a particular sender.

  Returns:
    List of Messages that match the criteria of the query. Note that the
    returned list contains Message IDs, you must use get with the
    appropriate ID to get the details of a Message.
  """
  try:
    response = service.users().messages().list(userId=user_id,
                                               q=query).execute()
    messages = []
    if 'messages' in response:
      messages.extend(response['messages'])

    while 'nextPageToken' in response:
      page_token = response['nextPageToken']
      response = service.users().messages().list(userId=user_id, q=query,
                                         pageToken=page_token).execute()
      messages.extend(response['messages'])

    return messages
  except(errors.HttpError, error):
    # Python 3 在except需加上 ()
    print('An error occurred: %s' % error)


def ListMessagesWithLabels(service, user_id, label_ids=[]):
  """List all Messages of the user's mailbox with label_ids applied.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    label_ids: Only return Messages with these labelIds applied.

  Returns:
    List of Messages that have all required Labels applied. Note that the
    returned list contains Message IDs, you must use get with the
    appropriate id to get the details of a Message.
  """
  try:
    response = service.users().messages().list(userId=user_id,
                                               labelIds=label_ids).execute()
    messages = []
    if 'messages' in response:
      messages.extend(response['messages'])

    while 'nextPageToken' in response:
      page_token = response['nextPageToken']
      response = service.users().messages().list(userId=user_id,
                                                 labelIds=label_ids,
                                                 pageToken=page_token).execute()
      messages.extend(response['messages'])

    return messages
  except(errors.HttpError, error):
    # Python 3 在except需加上 ()
    print('An error occurred: %s' % error)
"""Get Message with given ID.
"""

import base64
import email
from apiclient import errors

def GetMessage(service, user_id, msg_id):
  """Get a Message with given ID.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    msg_id: The ID of the Message required.

  Returns:
    A Message.
  """
  try:
    message = service.users().messages().get(userId=user_id, id=msg_id).execute()

    print('Message snippet: %s' % message['snippet'])

    return message
  except(errors.HttpError, error):
    print('An error occurred: %s' % error)


def GetMimeMessage(service, user_id, msg_id):
  """Get a Message and use it to create a MIME Message.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    msg_id: The ID of the Message required.

  Returns:
    A MIME Message, consisting of data from Message.
  """
  try:
    message = service.users().messages().get(userId=user_id, id=msg_id,
                                             format='raw').execute()

    print('Message snippet: %s' % message['snippet'])

    msg_str = base64.urlsafe_b64decode(message['raw'].encode('ASCII'))

    mime_msg = email.message_from_string(msg_str)

    return mime_msg
  except(errors.HttpError, error):
    print('An error occurred: %s' % error)
	
	
def main():
    """Shows basic usage of the Gmail API.

    Creates a Gmail API service object and outputs a list of label names
    of the user's Gmail account.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('gmail', 'v1', http=http)

    #for a_message in ListMessagesMatchingQuery(service, user_id='me', query='w'):
        #print(a_message['id'])
    mailid = "15385c2733b2ddb1"
    mes = GetMessage(service, user_id='me', msg_id=mailid)
    
    
    ''' file = open("檔案路徑" , "讀取方式" ,encoding = 'UTF-8')
		讀取方式:
		 'r' when the file will only be read, 
		 'w' for only writing (an existing file with the same name will be erased), and 
		 'a' opens the file for appending; any data written to the file is automatically added to the end. 
		 'r+' opens the file for both reading and writing. The mode argument is optional; 
		 'r' will be assumed if it’s omitted.
	'''

    w = csv.writer(open("C:\\Users\\CHICHI\\Desktop\\output.csv", "a", encoding='UTF-8'))
	#載入csv檔案，若沒有則建立
    for key, val in mes.items():
        w.writerow([key, val])
		#一行一行寫入
		
		#160404 1627 截止進度:下次要分析csv檔案 
	'''	理想的email 應該是下面這樣的模式
		{
		  "id": string,
		  "threadId": string,
		  "labelIds": [
			string
		  ],
		  "snippet": string,
		  "historyId": unsigned long,
		  "internalDate": long,
		  "payload": {
			"partId": string,
			"mimeType": string,
			"filename": string,
			"headers": [
			  {
				"name": string,
				"value": string
			  }
			],
			"body": users.messages.attachments Resource,
			"parts": [
			  (MessagePart)
			]
		  },
		  "sizeEstimate": integer,
		  "raw": bytes
		}
	'''	
		
		
		
    ''' 下面這幾行本來是要list labels用的 
    results = service.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])

    if not labels:
        print('No labels found.')
    else:
      print('Labels:')
      for label in labels:
        print(label['name'])
    '''

if __name__ == '__main__':
    main()