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
        $Path_log,
        [Parameter(Mandatory)]
        $DB_Mail,
        [Parameter(Mandatory)]
        $ArchiveDatabase
    )
    Set-JiraConfigServer -Server "$Jira"
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Mail_serv/PowerShell/ -Authentication Kerberos -Credential $Cred_Exch
    Import-PSSession $Session -AllowClobber -CommandName Enable-Mailbox,Get-GlobalAddressList,Update-GlobalAddressList,Get-OfflineAddressBook,Update-OfflineAddressBook,Get-AddressList,Update-AddressList -DisableNameChecking

    $LastName        = $UserAttr.LastName
    $FirstName       = $UserAttr.FirstName
    $SamAccountName  = $UserAttr.SamAccountName
    $PipeLine        = $UserAttr.PipeLine
    try{
        Add-Type -AssemblyName System.Web
        $length = Get-Random -Minimum 9 -Maximum 12
        $password = [System.Web.Security.Membership]::GeneratePassword($length, 2)
        $pass = ConvertTo-SecureString $password -AsPlainText -force
        $UserPrincipalName = $UserAttr.SamAccountName + "@" + $Domain
        $DisplayName = $UserAttr.LastName + " " + $UserAttr.FirstName
        $FullName = $UserAttr.LastName + " " + $UserAttr.FirstName + " " + $UserAttr.Initials

        $extension = @{
        'extensionAttribute1'= $UserAttr.Birth
        'extensionAttribute2' = $UserAttr.Initials
        'extensionAttribute5' = $UserAttr.Birth
        }
        $ADParameters = @{
        'Server'                = $DC
        'DisplayName'           = $DisplayName
        'GivenName'             = $UserAttr.FirstName
        'Office'                = $UserAttr.Office
        'Surname'               = $UserAttr.LastName
        'OfficePhone'           = $UserAttr.Phone
        'Department'            = $UserAttr.Department
        'Manager'               = $UserAttr.Manager
        'Title'                 = $UserAttr.JobTitle
        'UserPrincipalName'     = $UserPrincipalName
        'SamAccountName'        = $UserAttr.SamAccountName
        'Path'                  = $Path_OU
        'ChangePasswordAtLogon' = $true
        'Enabled'               = $true
        'AccountPassword'       = $pass
        'OtherAttributes'       = $extension
        }

        New-ADUser $FullName @ADParameters

        if ($UserAttr.Division -ne '') { Set-ADUser -Server $DC $SamAccountName -add @{extensionAttribute3 = "$UserAttr.Division" } }
        if ($UserAttr.Management -ne '') { Set-ADUser -Server $DC $SamAccountName -Replace @{extensionAttribute4 = "UserAttr.$Management" } }
    
        $user_creds = "Login: $SamAccountName`nPassword: $password"
        $parameters = @{
            Fields = @{
                customfield_10641 = "$user_creds"
            }
        }
        Set-JiraIssue @parameters -Issue "$PipeLine" -Credential $Cred_Jira
    
        Enable-Mailbox -identity $UserAttr.SamAccountName -Alias $UserAttr.SamAccountName -Database $DB_Mail -DomainController $DC
        Enable-Mailbox -identity $UserAttr.SamAccountName -archive -ArchiveDatabase $ArchiveDatabase -DomainController $DC
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