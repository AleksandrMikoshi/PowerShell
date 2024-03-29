# Server part

[Русский язык](https://github.com/AleksandrMikoshi/PowerShell/blob/main/Manage-Users/Server/Readme_ru.md)

The contents of this folder are located on the [JEA](https://learn.microsoft.com/en-us/powershell/scripting/learn/remoting/jea/overview?view=powershell-7.3) server to be able to control the actions of people who work with the form, as well as - there will be no need to grant elevated rights in Active Directory or Exchange to users, tk. in AD, all actions occur on behalf of [gMSA](https://learn.microsoft.com/en-us/windows-server/security/group-managed-service-accounts/group-managed-service-accounts-overview), and for Exchange and Jira, separate technical accounts are created, the password for which is known only to a limited number of people.   
[ManageUser.psm1](https://github.com/AleksandrMikoshi/PowerShell/blob/main/Manage-Users/Server/ManageUser.psm1) connects to the Passwork server via API and receives a login/password to connect to Jira

### You must specify

In the File [ManageUser.psm1](https://github.com/AleksandrMikoshi/PowerShell/blob/main/Manage-Users/Server/ManageUser.psm1) you must specify the variables:

| Variable              | What to specify |
|---|---|
| $URL_auth             | API address for connecting to Passwork            |
| $API_key_Passwork     | API key for authorization in Passwork             |
| $URL_data_passwords   | API address for getting passwords from Paswork    |
| $Case_Passwork        | Specifying an entry in Passwork to use            |
| $Domain               | Company domain (company.com)                      |
| $Jira                 | Jira Server URL                                   |