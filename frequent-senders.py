"""
Extract Message Header details from MBOX file
"""

import os, time
import mailbox
from email.parser import BytesParser
from email.policy import compat32
from collections import Counter

mailids = []


INBOX = 'C:/Users/WDAGUtilityAccount/AppData/Roaming/Thunderbird/Profiles/lmq6596d.default-release/ImapMail/imap.aol.com/INBOX'

print('Messages in ', INBOX)

mymail = mailbox.mbox(INBOX, factory=BytesParser(policy=compat32).parse)

for _, message in enumerate(mymail):
    sender = message['from']
    messageID = message['Message-ID']
    if(messageID == None): continue
    # print(sender)
    if sender is not None:
        try:
            domain = sender.rsplit("@", 1)[1]
            mailids.append(domain)
        except:
            continue
        

counter1 = Counter(mailids)

_elements = counter1.most_common()

for a in _elements:
    print(a)

