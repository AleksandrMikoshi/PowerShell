<#User account unlock feature#>
<#Функция разблокировки учетной записи пользователя#>
function Unlock_Account{
    param(
        [Parameter(Mandatory)]
        $User
    )
    $UserLogin = Get-ADUser -Filter "cn -eq '$User'" -Properties SamAccountName
    try{
        $ErrorActionPreference = 'Stop'

        $Color = "green"
        $Check = Get-ADUser -Identity $UserLogin.SamAccountName -Properties LockedOut
        if ($Check.LockedOut -eq $True) {
            Unlock-ADAccount -Identity $UserLogin.SamAccountName
            $Outcome="$User account unlocked"
        }
        else{
            $Outcome="$User account is not locked out"
        }
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