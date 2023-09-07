# Часть Client

Содержимое папки располагается на сервере который используется для RemoteApp приложений, для возможности отображать [Windows Form.NET](https://learn.microsoft.com/ru-ru/dotnet/desktop/winforms/overview/?view=netdesktop-7.0), так как подключение к сесии [JEA](https://learn.microsoft.com/en-us/powershell/scripting/learn/remoting/jea/overview?view=powershell-7.3) не позволяет отображать окна в виду того что подключение происходит посредством сессии в PowerShell  

### Необходимо самостоятельно указать

В файле [Start.ps1](https://github.com/AleksandrMikoshi/PowerShell/blob/main/Manage-Users/Client/Start.ps1) необходимо указать переменные:   

$Mail_serv = "Адрес почтового сервера для подключения по PowerShell"   
$Path_log = "Расположение файла в котором будут писаться логи в случае ошибок"   
$Path_OU = "Путь в Active Directory где хранятся учетные записи пользователей"   
$Path_fired = "Путь в Active Directory где хранятся учетные записи уволеных пользователей"   
$Path_OU_Outstaff = "Путь в Active Directory где хранятся учетные записи внешних пользователей"   
$Path_Mail = "Путь в Active Directory где хранятся учетные записи почтовых адресов"   
$Background_Image = "Путь на диске где располагается изображение для заднего фона приложения"   
$Icon = New-Object system.drawing.icon "Путь на диске где располагается изображение для иконки"   
$Label_Text = "Название компании либо любая другая подпись снизу формы"   
$Cat = "Путь на диске где располагается изображение для пункта 'Cat'"    