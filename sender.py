from exchangelib import Credentials, Configuration, Account, DELEGATE, Message, Mailbox, FileAttachment, HTMLBody
import requests

def attach_file(filename, message, id):
  with open(filename, 'rb') as f:
    file = FileAttachment(
    name=filename, content=f.read(),
    is_inline=True, content_id=id,
    )
  message.attach(file)