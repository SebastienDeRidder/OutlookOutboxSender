# OutlookOutboxSender
This script will send mails stuck in Outbox caused by Outlook addins

In some versions of Outlook, the sending of mails in Outbox after they have been viewed or edited in the Outbox doesn't work because of certain addins being enabled.
Delayed delivery of mails enables you to review or edit mails that are in the Outbox, waiting to be sent. Some addins cause Outlook to focus on the mail in the Outbox which removes the scheduled time because they are being edited.
If, for whatever reason, you are unable to disable the culprit addins (likely Adobe or Social Connector), you can use this script which will unfocus Outlook from the Outbox by moving the view to the Inbox and send all mails from the Outbox.

## usage
1. open the Visual Basic script editor in Outlook. ALT+F11
2. insert a new module
3. paste the code from sendFromOutbox.vbs into the module and save. CTRL+S
4. Run the script. ALT+F8
