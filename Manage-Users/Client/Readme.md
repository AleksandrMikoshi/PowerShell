# Client part

[Русский язык](https://github.com/AleksandrMikoshi/PowerShell/blob/main/Manage-Users/Client/Readme_ru.md)   

The contents of the folder are located on the server that is used for RemoteApp applications to be able to display [Windows Form.NET](https://learn.microsoft.com/en-us/dotnet/desktop/winforms/overview/?view=netdesktop-7.0) , since connecting to the [JEA](https://learn.microsoft.com/en-us/powershell/scripting/learn/remoting/jea/overview?view=powershell-7.3) session does not allow to display windows since connection occurs through a session in PowerShell   

To work correctly on the server with RemoteApp, you need to install the PowerShell module - [JiraPS](https://atlassianps.org/docs/JiraPS/)   

Next, depending on the settings of Jira fields, you need to specify your fields that contain the information necessary to create user accounts.

## You must specify yourself

In the file [Start.ps1] (https://github.com/AleksandrMikoshi/PowerShell/blob/main/Manage-Users/Client/Start.ps1) you need to specify the variables:

### Shared variables
| Variable              | What needs to be specified                                                                |
|---|---|
| $mail_serv            | "Mail server address for connection via PowerShell"                                       |
| $DB                   | "Mailbox database name"                                                                   |
| $path_log             | "File location where logs will be written in case of errors"                              |
| $Path_OU              | "Path in Active Directory where user accounts are stored"                                 |
| $path_fired           | "The path in Active Directory where terminated user accounts are stored"                  |
| $Path_OU_Outstaff     | "Path in Active Directory where external user accounts are stored"                        |
| $PathMail             | "The path in Active Directory where mail address accounts are stored"                     |
| $Background_Image     | "The path on the disk where the image for the background of the application is located"   |
| $ Icon                | New-Object system.drawing.icon "Disk path where icon image is located"                    |
| $Label_Text           | "Company name or any other signature at the bottom of the form"                           |
| $Cat                  | "Disk path where image for 'Cat' is located"                                              |

### Fields from Jira
Here is an example of my Jira fields, if necessary, you need to replace "Data_From_Jira.ps1" and "Data_From_Jira_Outstaff.ps1" with your own values in the functions

| Field             | What is in the box                                                            |
|---|---|
| customfield_10501 | This field contains the city where the employee (office) goes                 |
| customfield_10505 | This field contains the date of birth                                         |
| customfield_10708 | This field contains the manager                                               |
| customfield_10408 | This field contains the Last Name First Name and Patronymic of the employee   |
| customfield_10506 | This field contains the future login of the employee                          |
| customfield_10410 | This field contains the position of the employee                              |
| customfield_10502 | This field contains the department where the employee goes                    |
| customfield_10700 | This field contains the office where the employee goes (optional)             |
| customfield_10503 | This field contains the employee's phone number                               |
| customfield_10409 | This field contains the department where the employee goes                    |
| key               | This field contains the application number                                    |