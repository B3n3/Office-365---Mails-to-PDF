Office 365 - Mails to PDF
=============================

This handy python script crawls through your Office 365 mailbox and converts all unread mails to a separate PDF.
Most of my colleagues have a dedicated folder of SBB receipts.  
The script will iterate through the folder you define and create a PDF invoice for each received mail, containing the amount paid and date as filename.


--------------------------------------------------------------


Requirements:
* [Azure App registration](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade), described [here](https://pypi.org/project/O365/#authentication) in more detail.
* [wkhtmltopdf](https://wkhtmltopdf.org/downloads.html), needed to convert HTML to PDF.
* Python libraries from [requirements.txt](requirements.txt).  
  Install with `pip install < requirements.txt`.  
  I'd recommend setting up a virtual environment for python.
  
How to use?
-----------

1. Perform the Azure App registration
2. Copy and paste the clientID and clientSecret to [credentials.py](credentials.py).  
3. Initially call the [authenticate.py](authenticate.py) script to set-up the connection / authenticate this app.
4. Download and istall wkhtmltopdf.
5. Update [mail_to_pdf.py](mail_to_pdf.py).
    6. Update with the wkhtmltopdf path.
    7. Set the mailbox folder name in case you have your mails in a dedicated SBB receipts folder.
    7. Set the output directory

Now you can simply run [mail_to_pdf.py](mail_to_pdf.py).