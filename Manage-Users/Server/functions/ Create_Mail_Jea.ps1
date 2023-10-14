function Create_Mail_Jea {
    param (
        [Parameter(Mandatory)]
        [hashtable]$UserAttr,
        [Parameter(Mandatory)]
        [string]$DC,
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

    try{
        Enable-Mailbox -identity $UserAttr.SamAccountName -Alias $UserAttr.SamAccountName -Database $DB_Mail -DomainController $DC
        Enable-Mailbox -identity $UserAttr.SamAccountName -archive -ArchiveDatabase $ArchiveDatabase -DomainController $DC
        Get-GlobalAddressList | Update-GlobalAddressList
        Get-OfflineAddressBook | Update-OfflineAddressBook
        Get-AddressList | Update-AddressList
        Remove-PSSession $Session

        $Color = "green"
        $Outcome = "✓ Creating a mailbox"
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
        $Outcome = '× Creating a mailbox. Contact your system administrator'
        $Total = @{
            Color = $Color
            Outcome = $Outcome
        }
        $Total
    }
} 
