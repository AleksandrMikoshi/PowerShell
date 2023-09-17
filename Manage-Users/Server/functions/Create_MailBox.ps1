<#Mailbox creation function#>
<#Функция создания почтового ящика#>
function Create_MailBox {
    param (
        [Parameter(Mandatory)]
        [hashtable]$MailAttr,
        [Parameter(Mandatory)]
        [string]$DC,
        [Parameter(Mandatory)]
        [string]$Path_Mail,
        [Parameter(Mandatory)]
        $Mail_serv,
        [Parameter(Mandatory)]
        $Path_log,
        [Parameter(Mandatory)]
        $DB_Mail
    )
    $Name = $MailAttr.Name
    $UserPrincipalName = $MailAttr.UserPrincipalName
    $leght = Get-Random -Minimum 8 -Maximum 12
    $generate = [System.Web.Security.Membership]::GeneratePassword($leght,2)
    $password_mail = ConvertTo-SecureString $generate -AsPlainText -Force

    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Mail_serv/PowerShell/ -Authentication Kerberos -Credential $Cred_Exch
    Import-PSSession $Session -AllowClobber -CommandName New-Mailbox -DisableNameChecking

    try {
        $ErrorActionPreference = 'Stop'

        New-Mailbox -Name $Name -UserPrincipalName $UserPrincipalName -Password $password_mail -DisplayName $Name -OrganizationalUnit $Path_Mail -Database DB_Mail -DomainController $DC
        $Color = "green"
        $Outcome="Почтовый ящик $Name ($UserPrincipalName) создан"
        $Total = @{
            Color = $Color
            Outcome = $Outcome
        }
        $Total 
        Remove-PSSession $Session

    }
    catch {
        $Data = Get-Date
        $Exception = ($Error[0].Exception).Message
        $InvocationInfo = ($Error[0].InvocationInfo).PositionMessage
        $Value = "$Data" + " " + "$Exception"+"
        " + "$InvocationInfo"
        Add-Content -Path $Path_log -Value $Value
        Remove-PSSession $Session
        $Color = "red" 
        $Outcome = 'Выполнение не удалось. Обратитесь к системному администратору'
        $Total = @{
            Color = $Color
            Outcome = $Outcome
        }
        $Total 
    }

}