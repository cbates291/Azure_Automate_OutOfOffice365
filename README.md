# Azure_Automate_OutOfOffice365
Allows you to automate setting an OOO for a mailbox each day and over the weekend.

# Usage
This is designed for use in Azure automation and uses the stored Automation credentials feature.

- Update Line 9 with the name of your credentials in Azure.
- Update Line 38 with the group objectID from Azure. This uses a group so you can add multiple mailboxes and control them all at once.
- Update Lines 33,36,42,45 as needed. These times are set in EST but must convert to UTC (done in the script) so if it shows 14:00:00, it is actually 6PM once the 4hr conversion happens.
