<#Function to create an account for an external employee#>
<#Функция создания учетной записи для внешнего сотрудника#>
function Create_Outstaff_User {
    param (
        [Parameter(Mandatory)]
        [hashtable]$UserAttr,
        [Parameter(Mandatory)]
        [string]$DC,
        [Parameter(Mandatory)]
        [string]$Path_OU_Outstaff,
        [Parameter(Mandatory)]
        $Path_log
    ) 
    $LastName        = $UserAttr.LastName
    $FirstName       = $UserAttr.FirstName
    $Initials        = $UserAttr.Initials
    $SamAccountName  = $UserAttr.SamAccountName
    $Office          = $UserAttr.Office
    $Department      = $UserAttr.Department
    $Phone           = $UserAttr.Phone
    $Company         = $UserAttr.Company
    try{
        $ErrorActionPreference = 'Stop'
        $UserPrincipalName = $SamAccountName + "@" + $Domain
        $FullName = $LastName + " " + $FirstName + " " + $Initials
        $DisplayName = $LastName + " " + $FirstName
        $length = Get-Random -Minimum 8 -Maximum 11
        $password = [System.Web.Security.Membership]::GeneratePassword($length, 2)
        New-ADUser -Server $DC -Name $FullName -DisplayName $DisplayName -GivenName $FirstName -Office $Office `
        -Surname $LastName -OfficePhone $Phone -Department $Department -Manager $Manager -UserPrincipalName $UserPrincipalName `
        -SamAccountName $SamAccountName -Path $Path_OU_Outstaff -AccountPassword (ConvertTo-SecureString $password -AsPlainText -force) `
        -ChangePasswordAtLogon $true -Enabled $true
        Set-ADUser -Server $DC $SamAccountName -add @{company = $Company}
        Set-ADUser -Server $DC $SamAccountName -add @{extensionAttribute2 = $Initials }

        Add-ADGroupMember jira_users -Server $DC $SamAccountName
        Add-ADGroupMember Confluence_Users -Server $DC $SamAccountName
        Add-ADGroupMember owncloud_users -Server $DC $SamAccountName
    
        $user_creds = "Логин: $SamAccountName`nПароль: $password"
        $parameters = @{
            Fields = @{
                customfield_10641 = "$user_creds"
            }
        }
        Set-JiraIssue @parameters -Issue "$PipeLine" -Credential $cred
        Remove-PSSession $Session
        $Color = "green"
        $Outcome="User $LastName $FirstName created"
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
 
