# Часть Server

Содержимое данной папки располагается на сервере [JEA](https://learn.microsoft.com/en-us/powershell/scripting/learn/remoting/jea/overview?view=powershell-7.3) для возможности контролировать действия людей, которые работают с формой, а так же - не будет необходимости предоставлять повышенные права в Active Directory или Exchange пользователям, т.к. в AD все действия происходят от имени [gMSA](https://learn.microsoft.com/ru-ru/windows-server/security/group-managed-service-accounts/group-managed-service-accounts-overview), а для Exchange и Jira создаются отдельные технические учетные записи, пароль от которых известен только ограниченному числу лиц.   
[ManageUser.psm1](https://github.com/AleksandrMikoshi/PowerShell/blob/main/Manage-Users/Server/ManageUser.psm1) подключается к серверу Passwork по API и получает логин/пароль для подключения к Jira

### Необходимо самостоятельно указать

В Файле [ManageUser.psm1](https://github.com/AleksandrMikoshi/PowerShell/blob/main/Manage-Users/Server/ManageUser.psm1) необходимо указать переменные:   

$URL_auth = API адрес для подключения к Passwork   
$API_key_Passwork = API ключ для авторизации в Passwork   
$URL_data_passwords = API адрес для получения паролей из Paswork   
$Case_Passwork = Указание записи в Passwork которую необходимо использовать   
$Domain = Домен компании (company.com)   
$Jira = URL-адрес сервера Jira
