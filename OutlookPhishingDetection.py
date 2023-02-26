import win32com.client as client
outlook = client.Dispatch('Outlook.Application')
namespace = outlook.GetNameSpace('MAPI')
account = namespace.Folders['s-spal@lwsd.org']
inbox = account.Folders['Inbox']
def isPhishing(message):
    if 'deal' in message.Body.lower():
        return True
    elif 'https' and '@' in message.Body.lower():
        return True
    elif 'bit.ly' in message.Body.lower():
        return True
    elif 'suspicious activity' in message.Body.lower():
        return True
    elif 'payment information' in message.Body.lower():
        return True 
    elif 'payment details' in message.Body.lower():
        return True
    elif 'http://' in message.Body.lower():
        return True
    elif 'social security' in message.Body.lower():
        return True
    elif 'valued customer' in message.Body.lower():
        return True
    elif 'Sir' in message.Body.lower():
        return True    
    else:
        return False
junk_messages = [message for message in inbox.Items if isPhishing(message) == True]
junk_stuff = inbox.Folders['SusStuff']
for message in junk_messages:
    message.move(junk_stuff)
if len(junk_messages) != 0:
    print("Removed all suspicious emails")
else:
    print("No suspicious emails were found!!!")
if len(junk_messages) != 1:
    print("There were ", len(junk_messages), " messages")
else:
    print("There were ", len(junk_messages), " message")
