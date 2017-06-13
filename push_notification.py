#!/usr/bin/env python

import win32com.client
from pushbullet import Pushbullet

api_key = "PUSHBULLET_APIKEY"
proxy = {"https": "https://USER:PASSWORD@SERVER:PORT/"}

pb = Pushbullet(api_key, proxy=proxy)

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")


inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                    # the inbox. You can change that number to reference
                                    # any other folder
messages = inbox.Items
message = messages.GetFirst()
body_content = message.body
print(body_content)

push = pb.push_note(message.SenderName + ": " + message.subject, message.body)
