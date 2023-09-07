# Client part

[Русский язык](https://github.com/AleksandrMikoshi/PowerShell/blob/main/Manage-Users/Client/Readme_ru.md)

The contents of the folder are located on the server that is used for RemoteApp applications to be able to display [Windows Form.NET](https://learn.microsoft.com/en-us/dotnet/desktop/winforms/overview/?view=netdesktop-7.0) , since connecting to the [JEA](https://learn.microsoft.com/en-us/powershell/scripting/learn/remoting/jea/overview?view=powershell-7.3) session does not allow to display windows since connection occurs through a session in PowerShell

### You must specify

In the file [Start.ps1] (https://github.com/AleksandrMikoshi/PowerShell/blob/main/Manage-Users/Client/Start.ps1) you need to specify the variables:   

$Mail_serv = "Mail server address to connect via PowerShell"   
$Path_log = "File location where logs will be written in case of errors"   
$Path_OU = "The path in Active Directory where user accounts are stored"   
$Path_fired = "Active Directory path where fired user accounts are stored"   
$Path_OU_Outstaff = "The path in Active Directory where external user accounts are stored"   
$Path_Mail = "The path in Active Directory where mail address accounts are stored"   
$Background_Image = "Path on disk where the app's background image is located"   
$Icon = New-Object system.drawing.icon "Disk path where icon image is located"    
$Label_Text = "Company name or any other label at the bottom of the form"   
$Cat = "Disk path where image for item 'Cat' is located"