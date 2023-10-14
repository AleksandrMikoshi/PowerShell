function Create_User {
    param (
        [Parameter(Mandatory)]
        [hashtable]$UserAttr,
        [Parameter(Mandatory)]
        [string]$DC,
        [Parameter(Mandatory)]
        [string]$Path_OU,
        [Parameter(Mandatory)]
        $Path_log
    )
    Set-JiraConfigServer -Server "$Jira"

    $LastName        = $UserAttr.LastName
    $FirstName       = $UserAttr.FirstName
    $SamAccountName  = $UserAttr.SamAccountName
    $PipeLine        = $UserAttr.PipeLine
    try{
        Add-Type -AssemblyName System.Web
        $length = Get-Random -Minimum 9 -Maximum 12
        $password = [System.Web.Security.Membership]::GeneratePassword($length, 2)
        $pass = ConvertTo-SecureString $password -AsPlainText -force
        $UserPrincipalName = $UserAttr.SamAccountName + $Domain
        $DisplayName = $UserAttr.LastName + " " + $UserAttr.FirstName
        $FullName = $UserAttr.LastName + " " + $UserAttr.FirstName + " " + $UserAttr.Initials

        $extension = @{
        'extensionAttribute1'= $UserAttr.Birth
        'extensionAttribute2' = $UserAttr.Initials
        'extensionAttribute5' = $UserAttr.Birth
        }
        if ($UserAttr.Division -ne '') { $extension += @{extensionAttribute3 = "$UserAttr.Division" } }
        if ($UserAttr.Management -ne '') { $extension += @{extensionAttribute4 = "$UserAttr.Management" } }
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

        $user_creds = "Login: $SamAccountName`nPassword: $password"
        $parameters = @{
            Fields = @{
                customfield_10641 = "$user_creds"
            }
        }
        Set-JiraIssue @parameters -Issue "$PipeLine" -Credential $Cred_Jira
    
        $Color = "green"
        $Outcome = "✓ Create an account"
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
        $Outcome = '× Create an account. Contact your system administrator'
        $Total = @{
            Color = $Color
            Outcome = $Outcome
        }
        $Total
    }
} 