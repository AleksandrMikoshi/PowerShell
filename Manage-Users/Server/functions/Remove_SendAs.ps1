<#The function of disabling "Send As" rights for the mailbox#>
<#Функция отключения прав "Отправить как" для почтового ящика#>
function Remove_SendAs{
    param (
        [Parameter(Mandatory)]
        $Mail,
        [Parameter(Mandatory)]
        $Users,
        [Parameter(Mandatory)]
        $Mail_serv,
        [Parameter(Mandatory)]
        $Path_log,
        [Parameter(Mandatory)]
        $URL_auth
    )
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Mail_serv/PowerShell/ -Authentication Kerberos -Credential $Cred_Exch
    Import-PSSession $Session -AllowClobber -CommandName Get-Mailbox,Remove-ADPermission -DisableNameChecking

    try{
        $ErrorActionPreference = 'Stop'

        foreach ($User in $Users){
            Get-Mailbox -Identity $Mail | Remove-ADPermission -User $User -AccessRights ExtendedRight -ExtendedRights 'Send As' -Confirm:$false
        }
        Remove-PSSession $Session
        $Color = "green"
        $Outcome="'Send as' rights granted"
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