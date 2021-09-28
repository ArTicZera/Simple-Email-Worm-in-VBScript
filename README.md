# Simple-Email-Worm-in-VBScript
I made a simple Email-Worm in VBScript that spreads via email, using "MAPI" API from Microsoft Outlook. It works on Microsoft Outlook 2000 (idk if 95/98 works, I tested in 2000 and it works perfectly)
# How-it-works?
First, it copies itself to System Folder Path. Then it uses Microsoft Outlook API (MAPI) to get access to Address Book and all contacts.
After that, it sends an email to all contacts in the Outlook address book, which of course contains an attachment with the worm that has copied itself to the System
