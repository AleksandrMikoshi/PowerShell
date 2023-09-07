<#Function to disable all groups for a user account except for the "Domain Users" group#>
<#Функция для отключения всех групп у учетной записи пользователя кроме группы "Domain Users"#>
function Remove_Group_Jea{
    param(
        [Parameter(Mandatory)]
        $User
    )
    $UserLogin = Get-ADUser -Filter "cn -eq '$User'" -Properties SamAccountName

    try{
        $ErrorActionPreference = 'Stop'

        $group_Member = Get-ADPrincipalGroupMembership -Identity $UserLogin | Where-Object { $_.Name -ne "Domain Users" }
        Remove-ADPrincipalGroupMembership -Identity $UserLogin -MemberOf $group_Member -Confirm:$false

        $Color = "green"
        $Outcome="$User's access groups removed"
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

        $Color = "red"
        $Outcome = 'Execution failed. Contact your system administrator'
        $Total = @{
            Color = $Color
            Outcome = $Outcome
        }
        $Total 
    }
}