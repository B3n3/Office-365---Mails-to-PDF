import re

import pdfkit
from O365 import Account

import credentials

WKHTMLTOPDF_BINARY = 'wkhtmltopdf\\bin\\wkhtmltopdf.exe'
MAILBOX_FOLDER = 'SBB_Tickets'
OUTPUT_DIR = 'out/'

credentials = (credentials.clientID, credentials.clientSecret)
account = Account(credentials)

ticketFolder = account.mailbox().get_folder(folder_name=MAILBOX_FOLDER)
config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_BINARY)

i = 1
amountRe = re.compile('<b>CHF&nbsp;(.*)</b>')

for message in ticketFolder.get_messages():
    if (not message.is_read):
        amount = '0.00'
        expenseDate = message.created.strftime('%d.%m.%Y')

        try:
            amount = amountRe.search(message.body).group(1)
        except:
            pass

        filename = amount + ', ' + expenseDate + '.pdf'
        print(i, filename)

        try:
            pdfkit.from_string(message.body, OUTPUT_DIR + filename, configuration=config)
        except:
            pass
        i = i + 1
