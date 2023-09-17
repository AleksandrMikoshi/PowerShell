<#The function of disabling the user account, moving it to OU=Fired_User and disabling the mailbox#>
<#Функция отключения пользовательской учетной записи, перенос её в OU=Fired_User и отключения почтового ящика#>
function Fired_User_Jea{
    param(
        [Parameter(Mandatory)]
        $User,
        [Parameter(Mandatory)]
        $Mail_serv,
        [Parameter(Mandatory)]
        $DC,
        [Parameter(Mandatory)]
        $Path_log,
        [Parameter(Mandatory)]
        $Path_fired
    )
    $UserLogin = Get-ADUser -Filter "cn -eq '$User'" -Properties SamAccountName
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Mail_serv/PowerShell/ -Authentication Kerberos -Credential $Cred_Exch
    Import-PSSession $Session -AllowClobber -CommandName Set-Mailbox, Disable-mailbox, Get-GlobalAddressList, Update-GlobalAddressList, Get-OfflineAddressBook, Update-OfflineAddressBook, Get-AddressList, Update-AddressList -DisableNameChecking

    try{
        $ErrorActionPreference = 'Stop'
        $group_Member = Get-ADPrincipalGroupMembership -Identity $UserLogin | Where-Object { $_.Name -ne "Domain Users" }
        Remove-ADPrincipalGroupMembership -Identity $UserLogin -MemberOf $group_Member -Confirm:$false
    }
    catch {
        $Data = Get-Date
        $Exception = ($Error[0].Exception).Message
        $InvocationInfo = ($Error[0].InvocationInfo).PositionMessage
        $Value = "$Data" + " " + "$Exception"+"
        " + "$InvocationInfo"
        Add-Content -Path Path_log -Value $Value
    }
    finally{
        Set-Mailbox -Identity $UserLogin -HiddenFromAddressListsEnabled $true
        Set-ADUser -Identity $UserLogin -Enabled $false
        Set-ADUser -Identity $UserLogin -Manager $null
        Get-ADUser -Identity $UserLogin -server $DC | Move-ADObject -targetpath $Path_fired
  
        Disable-mailbox -Identity $UserLogin -Confirm:$false
        Get-GlobalAddressList | Update-GlobalAddressList
        Get-OfflineAddressBook | Update-OfflineAddressBook
        Get-AddressList | Update-AddressList
        Remove-PSSession $Session
        $Color = "green"
        $Outcome="Учётная запись $User отключена"
        $Total = @{
            Color = $Color
            Outcome = $Outcome
        }
        $Total 
    }
}