import win32com.client as win32


class EmailSender:
    def __init__(self):
        self.outlook_app = win32.Dispatch('Outlook.Application')

    def send_email(self, data, signature):
        to_address = data.get('email_to', None)
        subject = data.get('subject', 'Urgent Request')
        signature = signature
        body = data.get('body', 'No body') # body in the Excel or modify this line.
        body = body.format(name=data.get('name', ''), links=data.get('links', ''))
        body = f"{body}\n\n{signature}"
        cc = data.get('email_cc', None)

        mail = self.outlook_app.CreateItem(0)
        mail.To = to_address
        mail.CC = cc
        # mail.SentOnBehalfOfName = '' # uncomment and add on behalf name.
        mail.Subject = subject
        mail.Body = body
        mail.Send()
        print(f'Email to {to_address} sent successfully')

