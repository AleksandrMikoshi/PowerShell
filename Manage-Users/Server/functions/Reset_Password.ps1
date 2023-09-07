<#Function to reset user password#>
<#Функция для сброса пароля пользователю#>
function Reset_Password{
    param(
        [Parameter(Mandatory)]
        $User
    )
    $UserLogin = Get-ADUser -Filter "cn -eq '$User'" -Properties SamAccountName
    try{
        $ErrorActionPreference = 'Stop'

        Add-Type -AssemblyName System.Web
        $length = Get-Random -Minimum 9 -Maximum 12
        $Generate = [System.Web.Security.Membership]::GeneratePassword($length, 2)
        $NewPasswd = ConvertTo-SecureString $Generate -AsPlainText -force
        Set-ADAccountPassword $UserLogin -NewPassword $NewPasswd -Reset -PassThru | Set-ADuser -ChangePasswordAtLogon $True
                    
        $Color = "green"
        $Outcome="$Generate"

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