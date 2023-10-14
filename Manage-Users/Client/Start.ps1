<#Creating the necessary variables for further work, loading additional functions from the "functions" folder
Connecting to the server where the JEA service is deployed and importing functions from it#>
<#Создание необходимых переменных для дальнейшей работы, загрузка дополнительных функций из папки "functions"
Подключение к серверу где развернута служба JEA и импортирование функций с него#>
Add-Type -assembly System.Windows.Forms
 
[string]$DC = (Get-ADDomainController -Writable -Service GlobalCatalog -Discover).HostName
$JEA_Serv = 'JeaServer'
$JEA_Conf = 'NameJeaConfig'
$Mail_serv = 'MAIL_SERVER'
$Path_log = 'C:\JEA\Logs\Error.txt' #example
$Path_OU = 'OU=Users,DC=Company,DC=Com' #example
$Path_fired = 'OU=Fired_users,DC=Company,DC=Com' #example
$Path_OU_Outstaff = 'OU=Outstaff,DC=Company,DC=Com' #example
$Path_Mail = 'OU=Mail,DC=Company,DC=Com' #example
$Background_Image = "C:\Image\Background.jpg" #example
$Icon = New-Object system.drawing.icon "C:\Image\icon.ico" #example
$Label_Text = "Company_Name"
$Fields = 'customfield_10409,customfield_10708,customfield_10408,customfield_10506,customfield_10410,customfield_10502,customfield_10503,customfield_10505,customfield_10501,customfield_10700,Key'
$Fields_Outstaff = 'customfield_10708,customfield_10408,customfield_10506,customfield_10502,customfield_10503,customfield_11305,customfield_10501'

 
$FunctionFiles = Get-ChildItem -Path $PSScriptRoot\functions -Filter *.ps1
foreach ($file in $FunctionFiles) {
    . $file.FullName
}
 
$Session_JEA = New-PSSession -ConfigurationName $JEA_Conf -ComputerName $JEA_Serv
 Import-PSSession $Session_JEA -AllowClobber -Module ManageUser

main_menu 
