<#Group copy function#>
<#Функция копирования групп#>
function Copy_Group_Jea {
    param (
        [Parameter(Mandatory)]
        $SourceUser,
        [Parameter(Mandatory)]
        $TargetUser
    )
    $SourceUserLogin = Get-ADUser -Filter "cn -eq '$SourceUser'" -Properties SamAccountName
    $TargetUserLogin = Get-ADUser -Filter "cn -eq '$TargetUser'" -Properties SamAccountName

    try{
        $SourceGroups = Get-ADPrincipalGroupMembership -Identity $SourceUserLogin.SamAccountName | Where-Object {$_.name -ne "Domain Users"}
        Add-ADPrincipalGroupMembership -Identity $TargetUserLogin.SamAccountName -MemberOf $sourceGroups -ErrorAction Stop

        $Color = "green"
        $Outcome="Группы от пользователя $SourceUser успешно скопированы"
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
        $Outcome = 'Выполнение не удалось. Обратитесь к системному администратору'
        $Total = @{
            Color = $Color
            Outcome = $Outcome
        }
        $Total
    }
}