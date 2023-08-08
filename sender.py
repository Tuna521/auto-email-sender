from exchangelib import Credentials, Configuration, Account, DELEGATE, Message, Mailbox, FileAttachment, HTMLBody
import requests


def attach_file(filename, message, id):
	with open(filename, 'rb') as f:
		file = FileAttachment(
		name=filename, content=f.read(),
		is_inline=False, content_id=id,
		)
	message.attach(file)
 
 
def send_welcome_email(username, password):
	creds = Credentials(username, password)
	config = Configuration(service_endpoint='Outlook', credentials=creds)
	account = Account(primary_smtp_address='your email', credentials=creds, config=config, autodiscover=False, access_type=DELEGATE) #TODO: your email change
	
	emails = []
	
	m = Message(
		account=account,
		folder=account.sent,
		subject="subject",
		to_recipient=emails,
	)
	
	# attach_file("filename.png", m, "id??")
	html=open("email.txt", "r").read()
	m.body = HTMLBody(html)
	m.send()
	# m.send_and_save()
	
send_welcome_email("artsoc@imperial.ac.uk", password=ARTSOC_PASSWORD)