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
        $Path_log
    )
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Mail_serv/PowerShell/ -Authentication Kerberos -Credential $Cred_Exch
    Import-PSSession $Session -AllowClobber -CommandName Enable-Mailbox,Get-GlobalAddressList,Update-GlobalAddressList,Get-OfflineAddressBook,Update-OfflineAddressBook,Get-AddressList,Update-AddressList -DisableNameChecking

    $UserLogin = Get-ADUser -Filter "cn -eq '$User'" -Properties SamAccountName
    try{
        $ErrorActionPreference = 'Stop'

        Enable-Mailbox -identity $User -Alias $UserLogin -Database $DB -DomainController $DC
        Enable-Mailbox -identity $User -archive -ArchiveDatabase ARCHIVE -DomainController $DC
        Get-GlobalAddressList | Update-GlobalAddressList
        Get-OfflineAddressBook | Update-OfflineAddressBook
        Get-AddressList | Update-AddressList

        Remove-PSSession $Session
        $Color = "green"
        $Outcome="Mailbox $User@$Domain created"
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
        $Outcome = 'Execution failed. Contact your system administrator'
        $Total = @{
            Color = $Color
            Outcome = $Outcome
        }
        $Total
    }
}
 
