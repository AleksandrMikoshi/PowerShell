<#Create variables to further connect to Jira and Passwork servers if needed
Reading the functions lying in the "functions" folder to activate them in the session#>
<#Создание переменных для дальнейшего подключения к серверам Jira и Passwork при необходимости
Считываение функций лежащих в папке "functions" для их активации в сессии#>
$Domain = "DOMAIN"
$Jira = 'Jira_Server'
$Group = 'Default groups required to add'
$URL_auth = "https://URL/api/v4/auth/login/"
$API_key_Passwork = "API_KEY"
$URL_data_passwords = "https://URL/api/v4/passwords/"
$Case_Passwork = "CASE"

$URL_data_Passwork = $URL_data_passwords+$Case_Passwork
$URL_auth_Passwork = $URL_auth+$API_key_Passwork
$Auth_Passwork = Invoke-RestMethod -Method Post -Uri "$URL_auth_Passwork" -ContentType "application/json"
$Token = $Auth_Passwork.data.token
$Header = @{"Passwork-Auth" = "$Token"}
$Data_Passwork = Invoke-RestMethod -Method Get -Uri "$URL_data_Passwork" -Headers $Header -ContentType "application/json"
$Login_Passwork = $Data_Passwork.data.login
$PassCrypt = $Data_Passwork.data.cryptedPassword
$Pass_Passwork = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($PassCrypt))
$Pass_Passwork = ConvertTo-SecureString $Pass_Passwork -AsPlainText -Force
$LogPas = New-Object System.Management.Automation.PSCredential ($Login_Passwork, $Pass_Passwork)
$SecurePasswd = "Password" | ConvertTo-SecureString -AsPlainText -Force
$UserName = "Login"
$Cred_Jira = New-Object System.Management.Automation.PSCredential ($UserName, $SecurePasswd) 

$Exch_Pass = @{
    Name = "Cred_Exch"
    Value = $LogPas
    Option = "ReadOnly"
    Scope = "Script"
    Force = $true
}
New-Variable @Exch_Pass
 
$Jira_Pass = @{
    Name = "Cred_Jira"
    Value = $Cred_Jira
    Option = "ReadOnly"
    Scope = "Script"
    Force = $true
}
New-Variable @Jira_Pass

$FunctionFiles = Get-ChildItem -Path $PSScriptRoot\functions -Filter *.ps1
foreach ($file in $FunctionFiles) {
. $file.FullName
}