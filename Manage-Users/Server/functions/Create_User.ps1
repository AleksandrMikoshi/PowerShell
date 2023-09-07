<#User and mailbox creation function#>
<#Функция создания пользователи и почтового ящика#>
function Create_User {
    param (
        [Parameter(Mandatory)]
        [hashtable]$UserAttr,
        [Parameter(Mandatory)]
        [string]$DC,
        [Parameter(Mandatory)]
        [string]$Path_OU,
        [Parameter(Mandatory)]
        $Mail_serv,
        [Parameter(Mandatory)]
        $Path_log
    )
    Set-JiraConfigServer -Server "$Jira"
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Mail_serv/PowerShell/ -Authentication Kerberos -Credential $Cred_Exch
    Import-PSSession $Session -AllowClobber -CommandName Enable-Mailbox,Get-GlobalAddressList,Update-GlobalAddressList,Get-OfflineAddressBook,Update-OfflineAddressBook,Get-AddressList,Update-AddressList -DisableNameChecking

    $LastName        = $UserAttr.LastName
    $FirstName       = $UserAttr.FirstName
    $Initials        = $UserAttr.Initials
    $SamAccountName  = $UserAttr.SamAccountName
    $Office          = $UserAttr.Office
    $JobTitle        = $UserAttr.JobTitle
    $Department      = $UserAttr.Department
    $Management      = $UserAttr.Management
    $Phone           = $UserAttr.Phone
    $Birth           = $UserAttr.Birth
    $Task            = $UserAttr.Task
    $PipeLine        = $UserAttr.PipeLine
    $Manager         = $UserAttr.Manager
    $Division        = $UserAttr.Division 
 
    try{
        $UserPrincipalName = $SamAccountName + "@" + $Domain
        $DisplayName = $LastName + " " + $FirstName
        $FullName = $LastName + " " + $FirstName + " " + $Initials
        Add-Type -AssemblyName System.Web
        $length = Get-Random -Minimum 9 -Maximum 12
        $password = [System.Web.Security.Membership]::GeneratePassword($length, 2)
        New-ADUser -Server $DC -Name $FullName -DisplayName $DisplayName -GivenName $FirstName `
        -Office $Office -Surname $LastName -OfficePhone $Phone -Department $Department `
        -Manager $Manager -Title $JobTitle -UserPrincipalName $UserPrincipalName `
        -SamAccountName $SamAccountName -Path $Path_OU -Enabled $true `
        -AccountPassword (ConvertTo-SecureString $password -AsPlainText -force) -ChangePasswordAtLogon $true

        Set-ADUser -Server $DC $SamAccountName -add @{extensionAttribute1 = $Birth }
        Set-ADUser -Server $DC $SamAccountName -add @{extensionAttribute2 = $Initials }
        Set-ADUser -Server $DC $SamAccountName -add @{extensionAttribute15 = "$Birth" }
 
        $default_g = 'Jira_Users','Confluence_Users','owncloud_users','bitrix_users','jrnlng','ra_users'
        $default_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
    
        $user_creds = "Login: $SamAccountName`nPassword: $password"
        $parameters = @{
            Fields = @{
                customfield_10641 = "$user_creds"
            }
        }
        Set-JiraIssue @parameters -Issue "$PipeLine" -Credential $Cred_Jira
    
        Enable-Mailbox -identity $SamAccountName -Alias $SamAccountName -Database DAG_DB02 -DomainController $DC
        Enable-Mailbox -identity $SamAccountName -archive -ArchiveDatabase ARCHIVE -DomainController $DC
        Get-GlobalAddressList | Update-GlobalAddressList
        Get-OfflineAddressBook | Update-OfflineAddressBook
        Get-AddressList | Update-AddressList
        Remove-PSSession $Session
        $Color = "green"
        $Outcome="
        User $LastName $FirstName created
        Mailbox '$SamAccountName@$Domain' has been created"
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
