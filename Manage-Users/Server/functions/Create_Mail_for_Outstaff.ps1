<#The function of creating a mailbox for an external employee#>
<#Функция создания почтового ящика для внешнего сотрудника#>
function Create_Mail_for_Outstaff{
    param(
        [Parameter(Mandatory)]
        $User,
        [Parameter(Mandatory)]
        $DC,
        [Parameter(Mandatory)]
        $Mail_serv,
        [Parameter(Mandatory)]
        $Path_log,
        [Parameter(Mandatory)]
        $DB_Mail,
        [Parameter(Mandatory)]
        $ArchiveDatabase
    )
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Mail_serv/PowerShell/ -Authentication Kerberos -Credential $Cred_Exch
    Import-PSSession $Session -AllowClobber -CommandName Enable-Mailbox,Get-GlobalAddressList,Update-GlobalAddressList,Get-OfflineAddressBook,Update-OfflineAddressBook,Get-AddressList,Update-AddressList -DisableNameChecking

    $UserLogin = Get-ADUser -Filter "cn -eq '$User'" -Properties SamAccountName
    try{
        $ErrorActionPreference = 'Stop'

        Enable-Mailbox -identity $UserLogin -Alias $User -Database $DB_Mail -DomainController $DC
        Enable-Mailbox -identity $UserLogin -archive -ArchiveDatabase $ArchiveDatabase -DomainController $DC
        Get-GlobalAddressList | Update-GlobalAddressList
        Get-OfflineAddressBook | Update-OfflineAddressBook
        Get-AddressList | Update-AddressList

        Remove-PSSession $Session
        $Color = "green"
        $Outcome="Почтовый ящик $User@m2.ru создан"
        $Total = @{
            Color = $Color
            Outcome = $Outcome
        }
        $Total
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