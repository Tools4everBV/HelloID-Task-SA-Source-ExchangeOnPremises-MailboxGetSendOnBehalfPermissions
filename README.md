# HelloID-Task-SA-Source-ExchangeOnPremises-MailboxGetSendOnBehalfPermissions

## Prerequisites
- [ ] Execute the cmdlet **Enable-PsRemoting** on the **Exchange server** to which you want to connect.
- [ ] Within **IIS**, under the **Exchange Back End site** for the **Powershell sub-site**, check that the authentication method **Windows Authentication** is **enabled**.
- [ ] Permissions to manage the Exchange objects, the default AD group **Organization Management** should suffice, but please change this accordingly.


## Description
This code snippet executes the following tasks:

1. Imports the ExchangeOnlineManagement module.
2. Define `$mailboxGuid` based on the `selectedMailbox` data source input `$datasource.selectedMailbox.Guid`
3. Creates a session to Exchange using Remote PowerShell.
4. List all recipients in Exchange Online with `SendOnBehalf` permissions to the mailbox using the cmdlet: [Get-Mailbox](https://learn.microsoft.com/en-us/powershell/module/exchange/get-mailbox?view=exchange-ps)
5. Return a hash table for each user account using the `Write-Output` cmdlet.

> To view an example of the data source output, please refer to the JSON code pasted below.

```json
{
  "selectedMailbox": {
    "Guid": "7d53a91f-dd9d-41b3-94fb-143bd2fc6854"
  }
}
```
